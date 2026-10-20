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
import time
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
    names = ("projects", "tasks", "resources", "milestones", "settings", "saved_reports", "audit_log", "calendar_events", "resource_assignments")
    snapshot = {}
    for name in names:
        rows = c.execute(f"SELECT * FROM {name}").fetchall()
        snapshot[name] = [dict(row) for row in rows]
    snapshot["saved_at"] = datetime.now().isoformat(timespec="seconds")
    snapshot["schema_version"] = 2
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
    tables = ("projects", "tasks", "resources", "milestones", "settings", "saved_reports", "audit_log", "calendar_events", "resource_assignments")
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
    """Recalculate leaf-task completion, hierarchy progress, and date-driven statuses."""
    today = date.today()
    with conn() as c:
        projects = c.execute("SELECT id, start_date, end_date, status FROM projects").fetchall()
        for project in projects:
            tasks = [dict(r) for r in c.execute(
                "SELECT id, parent_task_id, status, progress, due_date FROM tasks WHERE project_id=?", (project["id"],)
            ).fetchall()]
            by_id = {int(t["id"]): t for t in tasks}
            children = {}
            for t in tasks:
                parent = t.get("parent_task_id")
                if parent is not None and int(parent) in by_id:
                    children.setdefault(int(parent), []).append(t)
            # Work units are leaf tasks only; parent tasks summarize their descendants.
            leaf_tasks = [t for t in tasks if not children.get(int(t["id"]))]
            def is_complete(t):
                return str(t.get("status") or "").strip().lower() == "completed"
            for t in sorted(tasks, key=lambda x: int(x["id"]), reverse=True):
                kids = children.get(int(t["id"]), [])
                if kids:
                    leaf_descendants = []
                    stack = list(kids)
                    while stack:
                        node = stack.pop()
                        node_kids = children.get(int(node["id"]), [])
                        if node_kids: stack.extend(node_kids)
                        else: leaf_descendants.append(node)
                    total_child = len(leaf_descendants)
                    done_child = sum(is_complete(n) for n in leaf_descendants)
                    pct = round(done_child * 100 / total_child) if total_child else 0
                    if total_child and done_child == total_child:
                        stval = "Completed"
                    else:
                        due = None
                        try: due = date.fromisoformat(str(t.get("due_date") or "")[:10])
                        except Exception: pass
                        leaf_overdue = False; leaf_at_risk = False
                        for leaf in leaf_descendants:
                            if is_complete(leaf): continue
                            try: leaf_due = date.fromisoformat(str(leaf.get("due_date") or "")[:10])
                            except Exception: leaf_due = None
                            if leaf_due and leaf_due < today: leaf_overdue = True
                            elif leaf_due and (leaf_due-today).days <= 7: leaf_at_risk = True
                        if due and due < today: stval = "Overdue"
                        elif due and (due-today).days <= 7: stval = "At Risk"
                        elif leaf_overdue: stval = "Overdue"
                        elif leaf_at_risk: stval = "At Risk"
                        elif done_child: stval = "In Progress"
                        else: stval = "Not Started"
                    c.execute("UPDATE tasks SET progress=?, status=? WHERE id=?", (pct, stval, t["id"]))
                else:
                    due = None
                    try: due = date.fromisoformat(str(t.get("due_date") or "")[:10])
                    except Exception: pass
                    if is_complete(t):
                        pct, stval = 100, "Completed"
                    elif due and due < today:
                        pct, stval = 0, "Overdue"
                    elif due and (due - today).days <= 7:
                        pct, stval = 0, "At Risk"
                    elif str(t.get("status") or "").strip().lower() == "in progress":
                        pct, stval = 0, "In Progress"
                    else:
                        pct, stval = 0, "Not Started"
                    c.execute("UPDATE tasks SET progress=?, status=? WHERE id=?", (pct, stval, t["id"]))
            total = len(leaf_tasks)
            done = sum(is_complete(t) for t in leaf_tasks)
            progress = round(done * 100 / total) if total else 0
            end_date = None
            try: end_date = date.fromisoformat(str(project["end_date"] or "")[:10])
            except Exception: pass
            if total and done == total:
                status, health = "Completed", "Healthy"
            elif end_date and end_date < today:
                status, health = "Overdue", "Critical"
            elif end_date and (end_date - today).days <= 7:
                status, health = "At Risk", "At Risk"
            elif done or progress:
                status, health = "In Progress", "Healthy"
            else:
                status, health = "Not Started", "Healthy"
            c.execute("UPDATE projects SET tasks_done=?, tasks_total=?, progress=?, days_left=?, status=?, health=? WHERE id=?",
                      (done, total, progress, (end_date - today).days if end_date else None, status, health, project["id"]))
        c.commit()
    st.session_state["_vektor_rollup_date"] = today.isoformat()


def sql(q: str, p: tuple = ()) -> None:
    """Execute an app mutation, then persist it and recalculate dependent project metrics."""
    normalized = " ".join(q.strip().lower().split())
    match = re.search(r"\b(?:insert\s+into|update|delete\s+from)\s+([a-z_]+)", normalized)
    table = match.group(1) if match else ""
    is_entity_mutation = table in {"projects", "tasks", "resources", "milestones", "saved_reports", "settings", "audit_log", "calendar_events", "resource_assignments"}
    with conn() as c:
        c.execute(q, p)
        c.commit()
    if table in {"tasks", "projects", "milestones"}:
        _recalculate_project_rollups()
    if is_entity_mutation:
        _persist_snapshot()


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
                id INTEGER PRIMARY KEY, title TEXT NOT NULL, event_type TEXT, project_id INTEGER,
                task_id INTEGER, resources TEXT, start_date TEXT, end_date TEXT, start_time TEXT,
                end_time TEXT, allocation INTEGER DEFAULT 0, repeat_rule TEXT, description TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS resource_assignments(
                id INTEGER PRIMARY KEY, resource_id INTEGER NOT NULL, project_id INTEGER NOT NULL,
                allocation INTEGER DEFAULT 0, start_date TEXT, end_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, UNIQUE(resource_id,project_id)
            );
            """
        )

        # No project, task, resource, milestone, or report demo data is seeded in Python.
        # Real workspace records are loaded from the Excel workbook in Settings.
        default_settings = [
            ("workspace_name", "Vektor"),
            ("workspace_description", "Project management workspace for planning, tracking, and delivery."),
            ("theme", "HPE Green"),
            ("language", "English"),
            ("date_format", "MMM DD, YYYY"),
            ("default_landing", "Overview"),
            ("default_project_status", "Not Started"),
            ("project_template", "Standard Project"),
            ("auto_project_id", "true"),
            ("require_project_desc", "true"),
            ("allow_project_attachments", "true"),
            ("project_archiving", "false"),
            ("default_task_status", "Not Started"),
            ("default_task_priority", "Medium"),
            ("task_dependencies", "true"),
            ("allow_subtasks", "true"),
            ("recurring_tasks", "false"),
            ("require_assignee", "true"),
            ("task_attachments", "true"),
            ("default_milestone_status", "Not Started"),
            ("milestones_enabled", "true"),
            ("milestones_require_due", "true"),
            ("milestone_dependencies", "true"),
            ("milestone_notifications", "true"),
            ("notification_frequency", "Daily digest"),
            ("integration_Microsoft 365", "false"),
            ("integration_Slack", "false"),
            ("integration_Microsoft Teams", "false"),
            ("integration_OneDrive", "false"),
            ("audit_logs", "true"),
            ("data_export", "true"),
            ("automatic_backups", "true"),
            ("maintenance_mode", "false"),
            ("session_timeout", "8 hours"),
        ]
        c.executemany("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", default_settings)
        c.execute("CREATE TABLE IF NOT EXISTS calendar_events( id INTEGER PRIMARY KEY, title TEXT NOT NULL, event_type TEXT, project_id INTEGER, task_id INTEGER, resources TEXT, start_date TEXT, end_date TEXT, start_time TEXT, end_time TEXT, allocation INTEGER DEFAULT 0, repeat_rule TEXT, description TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP )")
        c.execute("CREATE TABLE IF NOT EXISTS resource_assignments(id INTEGER PRIMARY KEY, resource_id INTEGER NOT NULL, project_id INTEGER NOT NULL, allocation INTEGER DEFAULT 0, start_date TEXT, end_date TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP, UNIQUE(resource_id,project_id))")
        # Safe schema migrations for existing local databases.
        resource_cols = [r[1] for r in c.execute("PRAGMA table_info(resources)").fetchall()]
        if "skills" not in resource_cols:
            c.execute("ALTER TABLE resources ADD COLUMN skills TEXT")
        task_cols = [r[1] for r in c.execute("PRAGMA table_info(tasks)").fetchall()]
        for col, definition in (("parent_task_id", "INTEGER"), ("source_id", "TEXT"), ("completion_date", "TEXT"), ("task_type", "TEXT DEFAULT 'Task'"), ("tags", "TEXT"), ("estimated_effort", "REAL"), ("actual_effort", "REAL")):
            if col not in task_cols:
                c.execute(f"ALTER TABLE tasks ADD COLUMN {col} {definition}")
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


def _clean_cell(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")
    text_value = str(value).strip()
    return "" if text_value.lower() in {"nan", "nat", "none"} else text_value


def import_portfolio_workbook(uploaded_file) -> dict:
    """Replace operational data with rows from the uploaded Portfolio workbook."""
    required = {"ID", "Program", "Initiative / Milestone / Task", "Start Date", "End Date", "Owner", "Status"}
    try:
        portfolio = pd.read_excel(uploaded_file, sheet_name="Portfolio", header=0, engine="openpyxl")
        if not required.issubset(set(portfolio.columns)):
            uploaded_file.seek(0)
            portfolio = pd.read_excel(uploaded_file, sheet_name="Portfolio", header=2, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"Could not read the Portfolio sheet: {exc}") from exc
    missing = required.difference(set(portfolio.columns))
    if missing:
        raise ValueError("Portfolio sheet is missing required columns: " + ", ".join(sorted(missing)))
    portfolio = portfolio.loc[portfolio["Program"].notna() & portfolio["Initiative / Milestone / Task"].notna()].copy()
    portfolio["Program"] = portfolio["Program"].map(_clean_cell)
    portfolio["ID"] = portfolio["ID"].map(_clean_cell)
    portfolio["Initiative / Milestone / Task"] = portfolio["Initiative / Milestone / Task"].map(_clean_cell)
    portfolio = portfolio[(portfolio["Program"] != "") & (portfolio["Initiative / Milestone / Task"] != "")]
    if portfolio.empty:
        raise ValueError("The Portfolio sheet contains no project or task rows to import.")
    # Create project records from program labels, keeping all other workbook tabs untouched.
    program_names = list(dict.fromkeys(portfolio["Program"].tolist()))
    project_ids = {name: idx + 1 for idx, name in enumerate(program_names)}
    project_rows = []
    for name in program_names:
        group = portfolio[portfolio["Program"] == name]
        starts = pd.to_datetime(group["Start Date"], errors="coerce").dropna()
        ends = pd.to_datetime(group["End Date"], errors="coerce").dropna()
        owners = [ _clean_cell(x) for x in group["Owner"].tolist() if _clean_cell(x) ]
        team = "Portfolio"
        project_rows.append((project_ids[name], name, owners[0] if owners else "", team, "Portfolio", "Not Started", "Medium",
            starts.min().strftime("%Y-%m-%d") if not starts.empty else "",
            ends.max().strftime("%Y-%m-%d") if not ends.empty else "", 0, None, 0, 0, "", GREEN,
            f"Imported from portfolio: {name}", f"Imported from the Portfolio sheet in the source workbook.", "Healthy"))
    # IDs are used to infer parent/child task hierarchy from dotted portfolio identifiers.
    task_records = []
    source_to_task = {}
    raw_items = []
    for _, row in portfolio.iterrows():
        source_id = _clean_cell(row.get("ID"))
        raw_items.append({"source_id": source_id, "program": _clean_cell(row.get("Program")),
            "name": _clean_cell(row.get("Initiative / Milestone / Task")), "start": _clean_cell(row.get("Start Date")),
            "end": _clean_cell(row.get("End Date")), "owner": _clean_cell(row.get("Owner")),
            "status": _clean_cell(row.get("Automated Status")) or _clean_cell(row.get("Status")), "update": _clean_cell(row.get("Update")),
            "remarks": _clean_cell(row.get("Remarks")), "app_key": _clean_cell(row.get("App Key")),
            "parent_key": _clean_cell(row.get("Parent Key")), "task_type": _clean_cell(row.get("Task Type")) or "Task",
            "tags": _clean_cell(row.get("Tags")), "dependency": _clean_cell(row.get("Dependency")),
            "estimated_effort": _clean_cell(row.get("Estimated Effort")), "actual_effort": _clean_cell(row.get("Actual Effort")), "priority": _clean_cell(row.get("Priority")) or "Medium"})
    for idx, item in enumerate(raw_items, 1):
        item["task_id"] = idx
        if item["source_id"]: source_to_task[item["source_id"]] = idx
    # Prefer the closest prior dotted ID prefix within the same program, without linking duplicate IDs to themselves.
    for item in raw_items:
        sid = item["source_id"]
        parent_id = None
        explicit_parent_key = item.get("parent_key", "")
        if explicit_parent_key:
            candidates = [x for x in raw_items if x.get("app_key") == explicit_parent_key and x["program"] == item["program"] and x["task_id"] < item["task_id"]]
            if candidates: parent_id = candidates[-1]["task_id"]
        explicit_parent = _clean_cell(next((r.get("Parent ID") for r in portfolio.to_dict("records") if _clean_cell(r.get("ID")) == sid and _clean_cell(r.get("Initiative / Milestone / Task")) == item["name"]), "")) if sid else ""
        if parent_id is None and explicit_parent:
            candidates = [x for x in raw_items if x["source_id"] == explicit_parent and x["program"] == item["program"] and x["task_id"] < item["task_id"]]
            if candidates: parent_id = candidates[-1]["task_id"]
        if parent_id is None and sid and "." in sid:
            prefixes = [".".join(sid.split(".")[:n]) for n in range(len(sid.split("."))-1, 0, -1)]
            for prefix in prefixes:
                candidates = [x for x in raw_items if x["source_id"] == prefix and x["program"] == item["program"] and x["task_id"] < item["task_id"]]
                if candidates:
                    parent_id = candidates[-1]["task_id"]
                    break
        item["parent_task_id"] = parent_id
    # Avoid duplicate source IDs from becoming ambiguous parent links: actual row order and nearest prior parent are retained.
    for item in raw_items:
        proj_id = project_ids[item["program"]]
        st_raw = item["status"].strip().lower()
        if st_raw in {"completed", "complete", "done"}:
            status, progress = "Completed", 100
        elif st_raw in {"in progress", "in-progress", "ongoing"}:
            status, progress = "In Progress", 0
        else:
            status, progress = "Not Started", 0
        due = item["end"]
        try:
            due_date = date.fromisoformat(due[:10]) if due else None
            if status != "Completed" and due_date and due_date < date.today(): status = "Overdue"
            elif status != "Completed" and due_date and (due_date-date.today()).days <= 7: status = "At Risk"
        except Exception: pass
        priority = item.get("priority") or ("High" if any(x in item["name"].lower() for x in ("critical", "urgent")) else "Medium")
        try: estimated_effort = float(item.get("estimated_effort") or 0) or None
        except Exception: estimated_effort = None
        try: actual_effort = float(item.get("actual_effort") or 0) or None
        except Exception: actual_effort = None
        task_records.append((item["task_id"], item["name"], proj_id, "Portfolio", status, priority,
            item["start"] or None, item["end"] or None, item["owner"] or None, progress, item.get("dependency", ""), " | ".join(x for x in (item["update"], item["remarks"]) if x),
            item["parent_task_id"], item["source_id"] or None, None, item.get("task_type") or "Task", item.get("tags") or "", estimated_effort, actual_effort))
    resource_names = list(dict.fromkeys(x["owner"] for x in raw_items if x["owner"]))
    with conn() as c:
        # Replace only operational records; keep workspace settings and audit history.
        for table in ("calendar_events", "resource_assignments", "milestones", "tasks", "projects", "resources", "saved_reports"):
            c.execute(f"DELETE FROM {table}")
        c.executemany("INSERT INTO projects(id,name,owner,team,category,status,priority,start_date,end_date,progress,days_left,tasks_done,tasks_total,tags,bar_color,short_desc,description,health) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", project_rows)
        c.executemany("INSERT INTO tasks(id,name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id,source_id,completion_date,task_type,tags,estimated_effort,actual_effort) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", task_records)
        for idx, name in enumerate(resource_names, 1):
            c.execute("INSERT INTO resources(id,name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?,?)",
                (idx, name, "", "Portfolio", "Team Member", "Available", 100, "Imported from Portfolio Owner column"))
        c.commit()
    # Optional app-managed assignment sheet; it is not required by the supplied Portfolio source.
    try:
        uploaded_file.seek(0)
        assignment_df = pd.read_excel(uploaded_file, sheet_name="Resource Assignments", engine="openpyxl")
    except Exception:
        assignment_df = pd.DataFrame()
    if not assignment_df.empty and {"Resource", "Project", "Allocation (%)"}.issubset(set(assignment_df.columns)):
        with conn() as c:
            resource_map = {str(row["name"]): int(row["id"]) for row in c.execute("SELECT id,name FROM resources").fetchall()}
            project_map = {str(row["name"]): int(row["id"]) for row in c.execute("SELECT id,name FROM projects").fetchall()}
            for _, assignment in assignment_df.iterrows():
                resource_name = _clean_cell(assignment.get("Resource")); project_name = _clean_cell(assignment.get("Project"))
                if resource_name not in resource_map or project_name not in project_map: continue
                try: allocation = max(0, min(200, int(float(assignment.get("Allocation (%)") or 0))))
                except Exception: allocation = 0
                start_value = _clean_cell(assignment.get("Start Date")) or date.today().isoformat()
                end_value = _clean_cell(assignment.get("End Date")) or ""
                c.execute("INSERT INTO resource_assignments(resource_id,project_id,allocation,start_date,end_date) VALUES(?,?,?,?,?) ON CONFLICT(resource_id,project_id) DO UPDATE SET allocation=excluded.allocation,start_date=excluded.start_date,end_date=excluded.end_date",
                    (resource_map[resource_name], project_map[project_name], allocation, start_value, end_value))
            c.commit()
    _recalculate_project_rollups()
    _persist_snapshot()
    return {"projects": len(project_rows), "tasks": len(task_records), "resources": len(resource_names)}


def current_data_workbook() -> bytes:
    """Export live operational tables to a workbook that can be edited and re-imported."""
    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        for table in ("projects", "tasks", "resources", "milestones", "calendar_events"):
            df(f"SELECT * FROM {table}").to_excel(writer, sheet_name=table.title()[:31], index=False)
        assignments = df("SELECT r.name AS Resource,p.name AS Project,ra.allocation AS 'Allocation (%)',ra.start_date AS 'Start Date',ra.end_date AS 'End Date' FROM resource_assignments ra JOIN resources r ON r.id=ra.resource_id JOIN projects p ON p.id=ra.project_id ORDER BY r.name,p.name")
        assignments.to_excel(writer, sheet_name="Resource Assignments", index=False)
        # Add the original portfolio layout as an import-compatible sheet.
        tasks = df("SELECT t.*, p.name AS program FROM tasks t LEFT JOIN projects p ON p.id=t.project_id")
        portfolio = pd.DataFrame({
            "PC": [""] * len(tasks), "ID": tasks.get("source_id", pd.Series([""]*len(tasks))).fillna(""),
            "Program": tasks.get("program", pd.Series([""]*len(tasks))).fillna(""),
            "Initiative / Milestone / Task": tasks.get("name", pd.Series([""]*len(tasks))).fillna(""),
            "Start Date": tasks.get("start_date", pd.Series([""]*len(tasks))).fillna(""),
            "End Date": tasks.get("due_date", pd.Series([""]*len(tasks))).fillna(""),
            "Owner": tasks.get("assignee", pd.Series([""]*len(tasks))).fillna(""),
            "Status": tasks.get("status", pd.Series([""]*len(tasks))).fillna(""),
            "Past Due Date": [""] * len(tasks), "Update": [""] * len(tasks), "Remarks": tasks.get("notes", pd.Series([""]*len(tasks))).fillna(""),
            "App Key": tasks.get("id", pd.Series([""]*len(tasks))).apply(lambda v: str(int(v)) if pd.notna(v) and str(v).strip() else ""),
            "Parent Key": tasks.get("parent_task_id", pd.Series([""]*len(tasks))).apply(lambda v: str(int(v)) if pd.notna(v) and str(v).strip() else ""),
            "Task Type": tasks.get("task_type", pd.Series(["Task"]*len(tasks))).fillna("Task"),
            "Tags": tasks.get("tags", pd.Series([""]*len(tasks))).fillna(""),
            "Dependency": tasks.get("dependency", pd.Series([""]*len(tasks))).fillna(""),
            "Priority": tasks.get("priority", pd.Series(["Medium"]*len(tasks))).fillna("Medium"),
            "Estimated Effort": tasks.get("estimated_effort", pd.Series([""]*len(tasks))).fillna(""),
            "Actual Effort": tasks.get("actual_effort", pd.Series([""]*len(tasks))).fillna("")})
        portfolio.to_excel(writer, sheet_name="Portfolio", startrow=2, index=False)
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
                cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="053241")
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = max((len(str(cell.value or "")) for cell in list(col)[:300]), default=10)
                ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 42)
    return buff.getvalue()


if not st.session_state.get("_vektor_db_initialized", False):
    init_db()
    try:
        _mongo_restored = _restore_snapshot()
    except Exception as _mongo_startup_error:
        st.session_state["_vektor_mongo_connected"] = False
        st.session_state["_vektor_mongo_error"] = str(_mongo_startup_error)
    st.session_state["_vektor_db_initialized"] = True
if not st.session_state.get("_vektor_mongo_connected", False):
    _mongo_message = st.session_state.get("_vektor_mongo_error", "MongoDB is not configured. Add MONGODB_URI to Streamlit Secrets or the environment to enable durable shared storage. The app is currently using its local cache.")
    st.warning(f"MongoDB is not connected. {_mongo_message}")
else:
    # Refresh shared state at most every 15 seconds on an interaction, not on every rerun.
    _now_mono = time.monotonic()
    if _now_mono - float(st.session_state.get("_vektor_last_mongo_check", 0)) >= 15:
        try:
            _restore_snapshot()
        except Exception as _mongo_refresh_error:
            st.session_state["_vektor_mongo_error"] = str(_mongo_refresh_error)
        st.session_state["_vektor_last_mongo_check"] = _now_mono

# Recalculate once per day; entity mutations recalculate immediately inside sql().
if st.session_state.get("_vektor_rollup_date") != date.today().isoformat():
    try:
        _recalculate_project_rollups()
    except Exception as _rollup_error:
        st.warning(f"Automatic project/task recalculation could not complete: {_rollup_error}")

# Load Dataframes
P = df("SELECT * FROM projects ORDER BY id")
T = df("SELECT t.*, p.name AS project FROM tasks t LEFT JOIN projects p ON p.id=t.project_id ORDER BY t.due_date")
R = df("SELECT * FROM resources ORDER BY id")
M = df("SELECT m.*, p.name AS project FROM milestones m LEFT JOIN projects p ON p.id=m.project_id ORDER BY m.due_date")

projects_list = P.name.tolist()
pmap = dict(zip(P.name, P.id))
teams_list = sorted(set(R.team.dropna().tolist() + P.team.dropna().tolist())) or ["Portfolio", "IT", "Operations", "Marketing", "Design", "Others"]

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
/* Modal Dialog Width & Header Styling
   Hide Streamlit's duplicate top-left dialog title while preserving its close button.
   The in-content icon/title headers remain the visible modal headings. */
div[data-testid="stDialog"] > div > div {{
    max-width: 1020px !important;
    width: 95vw !important;
    border-radius: 14px !important;
    padding: 0 !important;
}}
div[data-testid="stDialog"] h2 {{
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
}}
div[data-testid="stDialog"] [data-testid="stMarkdownContainer"] > div:first-child {{
    margin-top: -4px;
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
    project_tasks = df("SELECT * FROM tasks WHERE project_id=?", (pid,))
    project_task_status = project_tasks["status"].fillna("Not Started").value_counts().to_dict() if not project_tasks.empty else {}
    project_milestones_view = M[M["project_id"] == pid].copy() if not M.empty else pd.DataFrame()

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
            st.markdown(
                clean_html(f"<div style='font-size:12px;line-height:2.1;padding-top:14px'><div style='display:flex;justify-content:space-between'><span><span style='color:{GREEN}'>●</span> Completed</span><b>{project_task_status.get('Completed',0)}</b></div><div style='display:flex;justify-content:space-between'><span><span style='color:{BLUE}'>●</span> In Progress</span><b>{project_task_status.get('In Progress',0)}</b></div><div style='display:flex;justify-content:space-between'><span><span style='color:{ORANGE}'>●</span> At Risk</span><b>{project_task_status.get('At Risk',0)}</b></div><div style='display:flex;justify-content:space-between'><span><span style='color:{RED}'>●</span> Overdue</span><b>{project_task_status.get('Overdue',0)}</b></div><div style='display:flex;justify-content:space-between'><span><span style='color:{GRAY}'>●</span> Not Started</span><b>{project_task_status.get('Not Started',0)}</b></div></div>"),
                unsafe_allow_html=True,
            )
    with c_right, st.container(border=True):
        section_title("⏳ Timeline", "Project milestones")
        if project_milestones_view.empty:
            st.info("No milestones have been imported for this project.")
        else:
            _m_sorted=project_milestones_view.copy(); _m_sorted["_due"]=pd.to_datetime(_m_sorted["due_date"],errors="coerce"); _m_sorted=_m_sorted.sort_values("_due")
            for _,_m in _m_sorted.iterrows():
                st.markdown(f"<div style='display:flex;justify-content:space-between;gap:10px;border-bottom:1px solid #E2E8F0;padding:7px 0'><span>{html.escape(str(_m['name']))}</span><span>{fmt_date(_m.get('due_date'))} · {badge_html(_m.get('status'))}</span></div>",unsafe_allow_html=True)

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
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Days Left</span><span style="color:#00875A;font-weight:700">{int(r.days_left) if pd.notna(r.days_left) else 0} days</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px;margin-top:4px"><span style="color:#64748B">Tags</span><span>{tags_html}</span></div>
                </div>
                <hr style="border-color:#EEF2F6;margin:12px 0">
                <div style="font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:6px">📄 Project Summary</div>
                <div style="font-size:11.5px;color:#64748B;line-height:1.55">
                    The {html.escape(r['name'])} project aims to deliver a modern, user-friendly corporate website with improved performance, enhanced user experience, and updated content. The project is currently on track with {pct}% completion. Key development activities are in progress and the team is aligned for the upcoming milestones.
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

    with b_right:
        with st.container(border=True):
            section_title("🎯 Key Milestones", "View All Milestones →")
            project_milestones = M[M["project_id"] == pid].copy() if not M.empty else pd.DataFrame()
            tr_html = "".join(
                f"<tr><td><b>{idx}</b></td><td><span style='color:{GREEN if str(m.status)=='Completed' else BLUE};margin-right:6px'>●</span>{html.escape(str(m['name']))}</td><td>{fmt_date(m.get('due_date'))}</td><td>{badge_html(m.get('status'))}</td></tr>"
                for idx, (_, m) in enumerate(project_milestones.iterrows(), 1)
            )
            if not tr_html: tr_html = "<tr><td colspan='4'>No milestones have been imported for this project.</td></tr>"
            st.markdown(
                f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>#</th><th>Milestone</th><th>Due Date</th><th>Status</th></tr></thead><tbody>{tr_html}</tbody></table></div>',
                unsafe_allow_html=True,
            )

        with st.container(border=True):
            section_title("🕒 Recent Activity", "View All Activity →")
            activity_df = df("SELECT created_at,event,details FROM audit_log WHERE item_id=? OR (item_type='Project' AND details LIKE ?) ORDER BY id DESC LIMIT 8", (pid, f"%{r['name']}%"))
            act_tr = "".join(f"<tr><td>{html.escape(str(a.created_at))}</td><td><span style='color:{BLUE};margin-right:6px'>●</span>{html.escape(str(a.event or 'Updated'))}: {html.escape(str(a.details or ''))}</td><td>Workspace user</td></tr>" for _,a in activity_df.iterrows())
            if not act_tr: act_tr = "<tr><td colspan='3'>No recorded activity for this project yet.</td></tr>"
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
    edit_project_id = int(r["id"]) if r is not None else None

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
        stats = ["Not Started", "In Progress", "At Risk", "Overdue", "Completed", "On Track", "Delayed"]
        cur_st = str(r.status) if r is not None and r.status in stats else "Not Started"
        with c2:
            p_stat = st.selectbox("Status (automated)", stats, index=stats.index(cur_st), disabled=True, help="Calculated from completed task/subtask progress and the project end date.")
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
            s_dt = date.fromisoformat(r.start_date) if r is not None and r.start_date else date.today()
            p_start = st.date_input("Start Date *", s_dt)
        with c6:
            e_dt = date.fromisoformat(str(r.end_date)[:10]) if r is not None and r.end_date else date.today() + timedelta(days=60)
            p_end = st.date_input("Target End Date *", e_dt)

        all_tags_pool = ["Web", "Marketing", "Corporate", "IT", "Development", "Mobile", "Customer", "Data", "Cloud", "Infrastructure", "Security", "Compliance", "AI", "Innovation", "Automation", "Operations", "HR", "Training", "Procurement", "Sustainability", "ESG"]
        cur_tags = [t.strip() for t in str(r.tags or "").split(",") if t.strip() in all_tags_pool]
        p_tags = st.multiselect("Tags", all_tags_pool, default=cur_tags)

    with top_r, st.container(border=True):
        section_title("📄 Project Progress (auto-calculated)")
        pct_val = int(r.progress) if r is not None else 0
        st.plotly_chart(render_project_progress_donut(pct_val, 185), use_container_width=True, config={"displayModeBar": False})
        project_tasks = T[T["project_id"] == edit_project_id] if edit_project_id is not None and not T.empty else pd.DataFrame()
        status_counts = project_tasks["status"].fillna("Not Started").value_counts().to_dict() if not project_tasks.empty else {}
        total_task_count = len(project_tasks)
        st.markdown(clean_html(f"<div style='font-size:12px;line-height:2.1;padding:4px 8px'><div style='display:flex;justify-content:space-between'><span>● Completed</span><b>{status_counts.get('Completed',0)} / {total_task_count}</b></div><div style='display:flex;justify-content:space-between'><span>● In Progress</span><b>{status_counts.get('In Progress',0)}</b></div><div style='display:flex;justify-content:space-between'><span>● At Risk</span><b>{status_counts.get('At Risk',0)}</b></div><div style='display:flex;justify-content:space-between'><span>● Overdue</span><b>{status_counts.get('Overdue',0)}</b></div></div>"),unsafe_allow_html=True)

    with st.container(border=True):
        section_title("📋 Project Timeline & Milestones")
        edit_project_id = int(r["id"]) if r is not None else None
        project_milestones = M[M["project_id"] == edit_project_id].copy() if edit_project_id is not None and not M.empty else pd.DataFrame()
        if project_milestones.empty:
            st.info("No milestones have been imported for this project.")
        else:
            _ms_total=len(project_milestones); _ms_done=int((project_milestones["status"].astype(str)=="Completed").sum())
            st.progress(_ms_done/_ms_total if _ms_total else 0.0, text=f"{_ms_done} of {_ms_total} milestones completed")
        rows_mk = "".join(f"<tr><td>{i}</td><td>{html.escape(str(m['name']))}</td><td>{fmt_date(m.get('due_date'))}</td><td>{badge_html(m.get('status'))}</td><td>Managed in Milestones</td></tr>" for i,(_,m) in enumerate(project_milestones.iterrows(),1))
        if not rows_mk: rows_mk = "<tr><td colspan='5'>No milestones have been imported for this project.</td></tr>"
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
            days_l = max(0, (p_end - date.today()).days)
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
    h_left,h_r1,h_r2=st.columns([2.3,0.9,1.1])
    with h_left:
        st.markdown(f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Project Overview</div><div style='font-size:13.5px;color:#64748B;margin-top:2px'>Live portfolio progress, task status, dates, and team delivery.</div>",unsafe_allow_html=True)
    with h_r1: ov_proj_filter=st.selectbox("Project",["All Projects"]+projects_list,key="ov_proj_sel",label_visibility="collapsed")
    with h_r2: ov_date_range=st.date_input("Date Range",(date.today(),date.today()+timedelta(days=60)),key="ov_date_sel",label_visibility="collapsed")
    ov_projects=P.copy(); ov_tasks=T.copy(); ov_resources=R.copy(); ov_milestones=M.copy()
    if ov_proj_filter!="All Projects":
        ov_projects=ov_projects[ov_projects["name"]==ov_proj_filter]
        ov_tasks=ov_tasks[ov_tasks["project"]==ov_proj_filter] if "project" in ov_tasks else ov_tasks.iloc[0:0]
        ov_milestones=ov_milestones[ov_milestones["project"]==ov_proj_filter] if "project" in ov_milestones else ov_milestones.iloc[0:0]
    if isinstance(ov_date_range,tuple) and len(ov_date_range)==2:
        for _dfname in ("ov_tasks","ov_milestones"):
            _frame=locals()[_dfname]
            _datecol="due_date"
            if _datecol in _frame:
                _dates=pd.to_datetime(_frame[_datecol],errors="coerce")
                _frame=_frame[_dates.isna()|((_dates.dt.date>=ov_date_range[0])&(_dates.dt.date<=ov_date_range[1]))]
            if _dfname=="ov_tasks": ov_tasks=_frame
            else: ov_milestones=_frame
    task_total=len(ov_tasks); task_counts=ov_tasks["status"].fillna("Not Started").value_counts().to_dict() if not ov_tasks.empty else {}
    completed=int(task_counts.get("Completed",0)); in_progress=int(task_counts.get("In Progress",0)); on_hold=int(task_counts.get("On Hold",0)); overdue=int(task_counts.get("Overdue",0))
    k1,k2,k3,k4,k5,k6=st.columns(6)
    cards=[(k1,"Total Projects",len(ov_projects),"Active in filter","❖","#DCFCE7",GREEN),(k2,"Total Tasks",task_total,"In date range","☑","#DCFCE7",GREEN),(k3,"Completed Tasks",completed,f"{round(completed*100/task_total) if task_total else 0}%", "◕","#DCFCE7",GREEN),(k4,"In Progress",in_progress,f"{round(in_progress*100/task_total) if task_total else 0}%","▶","#DBEAFE",BLUE),(k5,"At Risk",int(task_counts.get("At Risk",0)),"Date-driven","⚠","#FEF3C7",ORANGE),(k6,"Overdue",overdue,"Past end date","✕","#FEE2E2",RED)]
    for col,label,value,sub,ico,bg,fg in cards:
        with col: kpi_card_mockup(label,str(value),sub_text=sub,icon_html=ico,icon_bg=bg,icon_fg=fg)
    st.write("")
    r2_l,r2_r=st.columns([1.55,1])
    with r2_l,st.container(border=True):
        a,b=st.columns([3,1])
        with a: st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>▤ &nbsp; Project Progress</div>",unsafe_allow_html=True)
        with b:
            if st.button("View All Projects →",key="ov_to_proj",use_container_width=True): st.session_state.page="Projects"; st.rerun()
        if ov_projects.empty: st.info("No projects loaded. Upload a portfolio workbook in Settings.")
        else:
            table_rows=[]
            for _,pr in ov_projects.iterrows():
                table_rows.append(f"<tr><td><b>{html.escape(str(pr['name']))}</b></td><td>{fmt_date(pr.start_date)}</td><td>{fmt_date(pr.end_date)}</td><td>{progress_html(pr.progress,GREEN)}</td><td>{badge_html(pr.status)}</td><td style='text-align:center'>{int(pr.days_left) if pd.notna(pr.days_left) else '—'}</td></tr>")
            st.markdown(f"<div class='vk-table-wrap'><table class='vk-table'><thead><tr><th>Project</th><th>Start</th><th>End</th><th>Progress</th><th>Status</th><th>Days Left</th></tr></thead><tbody>{''.join(table_rows)}</tbody></table></div>",unsafe_allow_html=True)
    with r2_r,st.container(border=True):
        a,b=st.columns([2.4,1.1])
        with a: st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>📅 &nbsp; Task Status</div>",unsafe_allow_html=True)
        with b:
            if st.button("View Task List →",key="ov_to_tasks1",use_container_width=True): st.session_state.page="Task List"; st.rerun()
        if ov_tasks.empty: st.info("No tasks in this date range.")
        else: chart_status_donut(ov_tasks,"status",center_label="Tasks",height=230)
    r3_l,r3_r=st.columns([1.55,1])
    with r3_l,st.container(border=True):
        section_title("📅 Upcoming Milestones","Live data")
        if ov_milestones.empty: st.info("No milestones match this date range.")
        else:
            mm=ov_milestones.copy(); mm["_due"]=pd.to_datetime(mm["due_date"],errors="coerce"); mm=mm.sort_values("_due").head(8)
            rows=[]
            for _,m in mm.iterrows():
                try: days=(date.fromisoformat(str(m.due_date)[:10])-date.today()).days
                except Exception: days="—"
                rows.append(f"<tr><td><b>{html.escape(str(m['name']))}</b></td><td>{html.escape(str(m.get('project','—')))}</td><td>{fmt_date(m.get('due_date'))}</td><td>{days}</td><td>{badge_html(m.get('status'))}</td></tr>")
            st.markdown(f"<div class='vk-table-wrap'><table class='vk-table'><thead><tr><th>Milestone</th><th>Project</th><th>Due</th><th>Days Left</th><th>Status</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>",unsafe_allow_html=True)
    with r3_r,st.container(border=True):
        a,b=st.columns([2.4,1.1])
        with a: st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>📌 &nbsp; Task Priority</div>",unsafe_allow_html=True)
        with b:
            if st.button("View Task List →",key="ov_to_tasks2",use_container_width=True): st.session_state.page="Task List"; st.rerun()
        if ov_tasks.empty or "priority" not in ov_tasks: st.info("No task priority data.")
        else: chart_status_donut(ov_tasks,"priority",center_label="Priority",height=230)
    r4_l,r4_r=st.columns([1.15,1])
    with r4_l,st.container(border=True):
        a,b=st.columns([2.6,1.1])
        with a: st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>👥 &nbsp; Team Workload</div>",unsafe_allow_html=True)
        with b:
            if st.button("View Resources →",key="ov_to_res",use_container_width=True): st.session_state.page="Resources"; st.rerun()
        if ov_resources.empty: st.info("No resources loaded.")
        else:
            rows=[]
            for _,res in ov_resources.head(10).iterrows():
                assigned=ov_tasks[ov_tasks["assignee"].astype(str)==str(res["name"])] if not ov_tasks.empty else pd.DataFrame()
                comp=int((assigned["status"].astype(str)=="Completed").sum()) if not assigned.empty else 0
                prog=int((assigned["status"].astype(str)=="In Progress").sum()) if not assigned.empty else 0
                ovd=int((assigned["status"].astype(str)=="Overdue").sum()) if not assigned.empty else 0
                capacity=int(res.get("capacity") or 0)
                rows.append(f"<tr><td>{avatar_html(res['name'])}</td><td style='text-align:center'>{len(assigned)}</td><td style='text-align:center'>{prog}</td><td style='text-align:center'>{comp}</td><td style='text-align:center'>{ovd}</td><td>{progress_html(capacity,GREEN)}</td></tr>")
            st.markdown(f"<div class='vk-table-wrap'><table class='vk-table'><thead><tr><th>Team Member</th><th>Assigned</th><th>In Progress</th><th>Completed</th><th>Overdue</th><th>Capacity</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>",unsafe_allow_html=True)
    with r4_r,st.container(border=True):
        section_title("🕒 Recent Activity","Audit log")
        activity=df("SELECT created_at,event,item_type,details FROM audit_log ORDER BY id DESC LIMIT 8")
        if activity.empty: st.info("Activity will appear here as users create or update projects, tasks, resources, and events.")
        else:
            rows=[]
            for _,a in activity.iterrows(): rows.append(f"<tr><td>{fmt_date(str(a.created_at)[:10])}</td><td>{html.escape(str(a.event or 'Updated'))}</td><td>{html.escape(str(a.item_type or 'Record'))}</td><td>{html.escape(str(a.details or ''))}</td></tr>")
            st.markdown(f"<div class='vk-table-wrap'><table class='vk-table'><thead><tr><th>Date</th><th>Action</th><th>Type</th><th>Details</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>",unsafe_allow_html=True)


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
        p_date_range = st.date_input("Range", (date.today(), date.today() + timedelta(days=60)), key="p_date_box", label_visibility="collapsed")
    with h_b:
        if st.button("＋ New Project", type="primary", use_container_width=True, key="btn_new_proj_top"):
            st.session_state.active_modal = ("edit", None)
            st.rerun()

    st.write("")
    # 5 KPI Cards Row matching Projects screenshots
    total_p = len(P)
    project_status_counts=P["status"].fillna("Not Started").value_counts().to_dict() if not P.empty else {}
    k1,k2,k3,k4,k5=st.columns(5)
    status_cards=[(k1,"TOTAL PROJECTS",total_p,"#00875A","❖"),(k2,"COMPLETED",int(project_status_counts.get("Completed",0)),GREEN,"✓"),(k3,"IN PROGRESS",int(project_status_counts.get("In Progress",0)),BLUE,"▶"),(k4,"AT RISK",int(project_status_counts.get("At Risk",0)),ORANGE,"⚠"),(k5,"OVERDUE",int(project_status_counts.get("Overdue",project_status_counts.get("Delayed",0))),RED,"✕")]
    for col,label,value,color,icon_value in status_cards:
        with col: kpi_card_mockup(label,str(value),sub_text=f"{round(value*100/total_p) if total_p else 0}% of projects" if label!="TOTAL PROJECTS" else "Current portfolio",bar_pct=round(value*100/total_p) if total_p else 0,icon_html=icon_value,icon_bg="#F1F5F9",icon_fg=color,bar_color=color)

    st.write("")
    # Filter & View Mode Row (No radio buttons — uses tile buttons for List View / Card View)
    if "proj_view_mode" not in st.session_state:
        st.session_state.proj_view_mode = "Card View"

    all_managers = sorted(P.owner.dropna().unique().tolist())
    all_tags = sorted({t.strip() for raw in P.tags.dropna() for t in str(raw).split(",") if t.strip()})

    f1, f2, f3, f4, f_sp, f_clr, f_lst, f_crd = st.columns([0.95, 0.95, 0.95, 0.95, 0.7, 0.75, 0.68, 0.68])
    with f1:
        sel_stat = st.selectbox("Status", ["All Statuses", "Completed", "In Progress", "At Risk", "Overdue", "Not Started", "On Track"], key="pf_stat", label_visibility="collapsed")
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
    if isinstance(p_date_range, tuple) and len(p_date_range) == 2:
        _ps = pd.to_datetime(vp["start_date"], errors="coerce")
        _pe = pd.to_datetime(vp["end_date"], errors="coerce")
        vp = vp[(_pe.isna() | (_pe.dt.date >= p_date_range[0])) & (_ps.isna() | (_ps.dt.date <= p_date_range[1]))]
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

            # Pagination Footer
            st.markdown(f"<div style='display:flex;align-items:center;justify-content:space-between;padding-top:8px;font-size:12.5px;color:#475569'><div>Showing all {len(vp)} filtered projects</div><div>Use the filters and sort controls above to refine the list.</div></div>",unsafe_allow_html=True)

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
    gh_l,gh_p,gh_s,gh_t=st.columns([2.0,1.1,1.0,0.75])
    with gh_l:
        st.markdown(f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Gantt Chart</div><div style='font-size:13.5px;color:#64748B;margin-top:2px'>Live project and task schedule with automated progress and date-based risk.</div>",unsafe_allow_html=True)
    with gh_p: gantt_project=st.selectbox("Project",["All Projects"]+projects_list,key="gantt_project_filter")
    with gh_s: gantt_status=st.selectbox("Status",["All Statuses","Completed","In Progress","Not Started","At Risk","Overdue"],key="gantt_status_filter")
    with gh_t: gantt_scale=st.selectbox("Scale",["Month","Week","Day"],key="gantt_scale")
    if st.session_state.pop("gantt_reset_to_today", False): st.session_state.pop("gantt_date_top", None)
    gantt_date_range=st.date_input("Date Range",(date.today(),date.today()+timedelta(days=60)),key="gantt_date_top")
    bar1,bar2,bar3=st.columns([1,1,5])
    with bar1:
        if st.button("Today",key="gantt_today",use_container_width=True):
            st.session_state["gantt_reset_to_today"]=True; st.rerun()
    with bar2:
        st.download_button("Export Schedule",excel_bytes({"Projects":P,"Tasks":T}),file_name="vektor_gantt_schedule.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="gantt_export")
    gantt_rows=[]
    for _,pr in P.iterrows():
        try: ps=date.fromisoformat(str(pr.start_date)[:10]); pe=date.fromisoformat(str(pr.end_date)[:10])
        except Exception: continue
        gantt_rows.append({"Label":f"▰ {pr['name']}","Project":str(pr["name"]),"Type":"Project","Start":ps,"Finish":max(pe,ps),"Status":str(pr.get("status") or "Not Started"),"Progress":int(pr.get("progress") or 0),"Owner":str(pr.get("owner") or "")})
    for _,task in T.iterrows():
        try: ts=date.fromisoformat(str(task.start_date)[:10]) if task.start_date and str(task.start_date)!="nan" else date.fromisoformat(str(task.due_date)[:10]); te=date.fromisoformat(str(task.due_date)[:10])
        except Exception: continue
        if te < ts: ts,te=te,ts
        gantt_rows.append({"Label":f"　↳ {task['name']}","Project":str(task.get("project") or ""),"Type":"Task","Start":ts,"Finish":max(te,ts),"Status":str(task.get("status") or "Not Started"),"Progress":int(task.get("progress") or 0),"Owner":str(task.get("assignee") or "Unassigned")})
    gantt_df=pd.DataFrame(gantt_rows)
    if not gantt_df.empty:
        if gantt_project!="All Projects": gantt_df=gantt_df[gantt_df["Project"]==gantt_project]
        if gantt_status!="All Statuses": gantt_df=gantt_df[gantt_df["Status"]==gantt_status]
        if isinstance(gantt_date_range,tuple) and len(gantt_date_range)==2:
            gantt_df=gantt_df[(gantt_df["Finish"]>=gantt_date_range[0])&(gantt_df["Start"]<=gantt_date_range[1])]
    with st.container(border=True):
        if gantt_df.empty:
            st.info("No dated projects or tasks match the selected filters. Import or update data in Settings.")
        else:
            color_map={"Completed":GREEN,"In Progress":BLUE,"Not Started":GRAY,"At Risk":ORANGE,"Overdue":RED,"On Track":GREEN,"Delayed":RED}
            fig_g=px.timeline(gantt_df,x_start="Start",x_end="Finish",y="Label",color="Status",color_discrete_map=color_map,hover_data={"Project":True,"Type":True,"Progress":True,"Owner":True,"Start":True,"Finish":True,"Status":True})
            fig_g.update_yaxes(autorange="reversed",title=None,tickfont=dict(size=11))
            tickformat={"Month":"%b %Y","Week":"%d %b","Day":"%d %b"}.get(gantt_scale,"%b %Y")
            fig_g.update_xaxes(tickformat=tickformat,showgrid=True,gridcolor="#E2E8F0",title=None,rangeselector=None)
            fig_g.update_layout(height=max(360,len(gantt_df)*30),margin=dict(l=8,r=20,t=16,b=12),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",legend_title_text="Status",bargap=0.28)
            st.plotly_chart(fig_g,use_container_width=True,config={"displayModeBar":False})
            st.dataframe(gantt_df.sort_values(["Project","Start","Type"]),hide_index=True,use_container_width=True,height=260)


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
            proj = st.selectbox("Project *", projects_list or ["No projects loaded"], index=0, key="add_t_proj")
        with r1_b:
            parent_options = ["None (Top Level Task)"] + [f"{int(r.id)} · {r.name}" for _, r in T.iterrows()]
            parent_t = st.selectbox("Parent Task", parent_options, key="add_t_parent")

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
            t_stat = st.selectbox("Status *", ["🟦 Not Started", "🔵 In Progress", "🟢 Completed"], index=0, key="add_t_stat")

        r4_a, r4_b, r4_c = st.columns(3)
        with r4_a:
            t_start = st.date_input("Start Date *", date.today(), key="add_t_start")
        with r4_b:
            t_end = st.date_input("End Date *", date.today() + timedelta(days=6), key="add_t_end")
        with r4_c:
            st.number_input("Duration (Days)", min_value=1, max_value=365, value=7, key="add_t_dur")

        r5_a, r5_b, r5_c = st.columns([1.1, 1.1, 1])
        with r5_a:
            assignees_pool = R.name.dropna().astype(str).tolist() or ["Unassigned"]
            t_ass = st.selectbox("Assignee *", assignees_pool, index=0, key="add_t_ass")
        with r5_b:
            st.caption("Progress is calculated automatically from completed subtasks.")
        with r5_c:
            t_type = st.selectbox("Task Type", ["Task", "Milestone", "Subtask"], key="add_t_type")

        r6_a, r6_b, r6_c = st.columns([1.1, 1.1, 1])
        with r6_a:
            st.markdown("<div style='font-size:12.5px;font-weight:600;color:#0F172A;margin-bottom:6px'>Milestone</div>", unsafe_allow_html=True)
            t_is_milestone = st.toggle("◆ Mark as Milestone", value=False, key="add_t_ms_tog")
        with r6_b:
            dep_options = ["None"] + [f"{int(row.id)} · {row.name}" for _, row in T.iterrows()]
            t_dep = st.selectbox("Dependencies", dep_options, key="add_t_dep")
        with r6_c:
            t_tags = st.multiselect("Tags", ["Planning", "Design", "Development", "Testing", "Launch", "Research", "Strategy"], placeholder="Select tags", key="add_t_tags")

        r7_a, r7_b, r7_c = st.columns([1, 1, 1.2])
        with r7_a:
            t_est = st.number_input("Estimated Effort (Hours)", min_value=0.0, value=0.0, step=1.0, key="add_t_est")
        with r7_b:
            t_act = st.number_input("Actual Effort (Hours)", min_value=0.0, value=0.0, step=1.0, key="add_t_act")
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
                if t_end < t_start:
                    st.error("End date cannot be earlier than start date.")
                elif not t_name.strip():
                    st.error("Task Name is required.")
                else:
                    parent_id = None
                    if parent_t != "None (Top Level Task)":
                        try: parent_id = int(parent_t.split(" · ", 1)[0])
                        except Exception: parent_id = None
                    assignee = str(t_ass).split(" (", 1)[0]
                    dependency_value = "" if t_dep == "None" else t_dep
                    task_type_value = "Milestone" if t_is_milestone else t_type
                    tags_value = ",".join(t_tags)
                    sql("INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id,task_type,tags,estimated_effort,actual_effort) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (t_name.strip(), pmap.get(proj), "Portfolio", clean_stat, clean_prio, t_start.isoformat(), t_end.isoformat(), assignee,
                         0, dependency_value, t_notes or t_desc, parent_id, task_type_value, tags_value, float(t_est), float(t_act)))
                    new_id = int(df("SELECT MAX(id) AS id FROM tasks").iloc[0]["id"])
                    for subtask_name in st.session_state.get("add_t_sub_items", []):
                        # Checklist entries are represented as child task rows so completion drives progress.
                        sql("INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                            (subtask_name, pmap.get(proj), "Portfolio", "Not Started", clean_prio, t_start.isoformat(), t_end.isoformat(), assignee, 0, "", "Checklist subtask", new_id))
                    if task_type_value == "Milestone":
                        sql("INSERT INTO milestones(name,project_id,team,status,start_date,due_date,completion_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?)",
                            (t_name.strip(), pmap.get(proj), "Portfolio", clean_stat, t_start.isoformat(), t_end.isoformat(), date.today().isoformat() if clean_stat == "Completed" else "", dependency_value, t_notes or t_desc))
                    st.session_state["add_t_sub_items"] = []
                    log_event("Created", "Task", new_id, t_name.strip())
                    if not create_another:
                        st.session_state.pop("task_modal", None)
                    st.rerun()

    @st.dialog("View Task", width="large")
    def view_task_modal(task_rec: dict) -> None:
        tname = str(task_rec.get("name") or "Untitled Task")
        tproj = str(task_rec.get("project") or "Unassigned Project")
        tstart = str(task_rec.get("start_date") or "—")
        tend = str(task_rec.get("end_date") or "—")
        tdur = str(task_rec.get("duration") or "—")
        tpct = int(task_rec.get("pct", 0) or 0)
        tstat = str(task_rec.get("status") or "Not Started")
        tprio = str(task_rec.get("priority") or "Medium")
        tass = str(task_rec.get("assignee") or "Unassigned")
        task_id = int(task_rec.get("id") or 0)
        child_tasks = T[T["parent_task_id"].fillna(-1).astype(int) == task_id].copy() if task_id and "parent_task_id" in T else pd.DataFrame()
        child_count = len(child_tasks)
        child_done = int((child_tasks["status"].astype(str) == "Completed").sum()) if child_count else 0
        checklist_rows_html = "".join(f"<div style='display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9;padding:4px 0'><span>{'✅' if str(row.status)=='Completed' else '☐'} &nbsp;{html.escape(str(row['name']))}</span><span style='color:#64748B'>{html.escape(str(row.status))}</span></div>" for _,row in child_tasks.iterrows()) or "<div style='color:#64748B;padding:6px 0'>No subtasks recorded for this task.</div>"
        parent_id = task_rec.get("parent_task_id")
        parent_row = T[T["id"] == parent_id] if pd.notna(parent_id) and "id" in T else pd.DataFrame()
        parent_label = str(parent_row.iloc[0]["name"]) if not parent_row.empty else "Top-level task"
        task_description = str(task_rec.get("notes") or "No description recorded.")
        task_activity = df("SELECT created_at,event,details FROM audit_log WHERE item_id=? OR (item_type='Task' AND details LIKE ?) ORDER BY id DESC LIMIT 5", (task_id, f"%{tname}%"))
        activity_html = "".join(f"<div style='display:grid;grid-template-columns:150px 1fr;gap:8px;border-bottom:1px solid #F1F5F9;padding:5px 0'><span style='color:#64748B;font-size:11.5px'>{html.escape(str(a.created_at))}</span><span>{html.escape(str(a.event or 'Updated'))}: {html.escape(str(a.details or ''))}</span></div>" for _,a in task_activity.iterrows()) or "<div style='color:#64748B;padding:6px 0'>No recorded activity for this task yet.</div>"

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
                        {badge_html(tstat)}
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
                            {html.escape(task_description)}
                        </div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">👤 &nbsp;Assignee</div>
                        <div style="background:#F8FAFC;border:1px solid #EDF2F7;border-radius:8px;padding:14px;display:flex;align-items:center;gap:12px">
                            <div style="width:44px;height:44px;border-radius:50%;background:#0284C7;color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px">JM</div>
                            <div>
                                <div style="font-size:14px;font-weight:800;color:#0F172A">{html.escape(tass)}</div>
                                <div style="font-size:12px;color:#64748B">UX Designer</div>
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
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">☰ &nbsp;{html.escape(parent_label)}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Milestone</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600"><span style="color:#7C3AED">◆</span> &nbsp;Not a milestone</div>
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
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">💼 &nbsp;Task</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">🏷 &nbsp;Tags</div>
                        <div>{tag_pill_html('Design')} {tag_pill_html('Web')}</div>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:1.35fr 1fr;gap:12px;margin-bottom:12px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">☑ &nbsp;Checklist / Subtasks ({child_count})</div>
                        <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
                            <div style="flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden"><div style="width:{round(child_done*100/child_count) if child_count else 0}%;height:100%;background:#00875A"></div></div>
                            <span style="font-size:11.5px;color:#475569;font-weight:600">{child_done} of {child_count} completed</span>
                        </div>
                        <div style="font-size:12.5px;line-height:2.1">
{checklist_rows_html}
                        </div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">📎 &nbsp;Attachments</div>
                        <div style="border:1px solid #EDF2F7;border-radius:8px;padding:8px 10px;display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
                            <div style="display:flex;align-items:center;gap:10px">
                                <span style="background:#EF4444;color:white;font-size:10px;font-weight:700;padding:6px 6px;border-radius:6px">PDF</span>
                                <div><div style="font-size:12.5px;font-weight:700">No attachment stored</div><div style="font-size:11px;color:#64748B"></div></div>
                            </div>
                            <span style="color:#475569;font-weight:700">⬇ &nbsp;⋮</span>
                        </div>
                        <div style="border:1px solid #EDF2F7;border-radius:8px;padding:8px 10px;display:flex;align-items:center;justify-content:space-between">
                            <div style="display:flex;align-items:center;gap:10px">
                                <span style="background:#1473E6;color:white;font-size:10px;font-weight:700;padding:6px 6px;border-radius:6px">IMG</span>
                                <div><div style="font-size:12.5px;font-weight:700">No attachment stored</div><div style="font-size:11px;color:#64748B"></div></div>
                            </div>
                            <span style="color:#475569;font-weight:700">⬇ &nbsp;⋮</span>
                        </div>
                    </div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white;margin-bottom:14px">
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">💬 &nbsp;Recent Activity / Notes</div>
<div style="font-size:12.5px;line-height:2.2">{activity_html}</div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        b_sp, b_cl, b_ed = st.columns([3.8, 0.9, 1.1])
        with b_cl:
            if st.button("Close", key="close_v_task", use_container_width=True):
                st.session_state.pop("task_modal", None)
                st.rerun()
        with b_ed:
            if st.button("✎ Edit Task", type="primary", key="to_edit_v_task", use_container_width=True):
                for _key in list(st.session_state.keys()):
                    if _key.startswith(("ed_t_", "ed_sub_", "ed_child_done_", "ed_p_", "ed_dep_", "ed_est_", "ed_act_")):
                        st.session_state.pop(_key, None)
                st.session_state["ed_sub_items"] = []
                st.session_state.task_modal = ("edit", task_rec)
                st.rerun()

    @st.dialog("Edit Task", width="large")
    def edit_task_modal(task_rec: dict) -> None:
        tname = task_rec.get("name", "Untitled Task")
        tproj = task_rec.get("project", "Unassigned Project")
        twbs = task_rec.get("wbs", str(task_rec.get("id", "")))
        ed_task_id = int(task_rec.get("id") or 0)
        ed_children = T[T["parent_task_id"].fillna(-1).astype(int) == ed_task_id].copy() if ed_task_id and "parent_task_id" in T else pd.DataFrame()

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
                value=str(task_rec.get("notes") or ""),
                height=92,
                key="ed_t_desc",
            )
        with e2_r:
            ass_list = R.name.dropna().astype(str).tolist() or ["Unassigned"]
            current_assignee = str(task_rec.get("assignee") or "Unassigned")
            if current_assignee not in ass_list: ass_list = [current_assignee] + ass_list
            ed_ass = st.selectbox("👤 Assignee *", ass_list, index=ass_list.index(current_assignee), key="ed_t_ass")

        e3_a, e3_b, e3_c, e3_d = st.columns(4)
        with e3_a:
            ed_start_value = task_rec.get("start_iso") or task_rec.get("start_date")
            ed_start_default = date.fromisoformat(str(ed_start_value)[:10]) if ed_start_value and str(ed_start_value) not in {"—", "nan", "None"} else date.today()
            ed_start = st.date_input("Start Date *", ed_start_default, key="ed_t_start")
        with e3_b:
            ed_end_value = task_rec.get("due_iso") or task_rec.get("due_date") or task_rec.get("end_date")
            ed_end_default = date.fromisoformat(str(ed_end_value)[:10]) if ed_end_value and str(ed_end_value) not in {"—", "nan", "None"} else date.today() + timedelta(days=6)
            ed_end = st.date_input("End Date *", ed_end_default, key="ed_t_end")
        with e3_c:
            try: _duration_days=max(1,(date.fromisoformat(str(ed_end_value)[:10])-date.fromisoformat(str(ed_start_value)[:10])).days+1)
            except Exception: _duration_days=1
            st.metric("Duration (Days)", _duration_days)
        with e3_d:
            st.metric("Progress (automatic)", f"{int(task_rec.get('pct',0) or 0)}%")

        e4_a, e4_b, e4_c, e4_d = st.columns(4)
        with e4_a:
            ed_prio_opts = ["⚑ High", "⚑ Medium", "⚑ Low"]
            ed_prio_current = "⚑ " + str(task_rec.get("priority") or "Medium")
            ed_prio = st.selectbox("Priority *", ed_prio_opts, index=ed_prio_opts.index(ed_prio_current) if ed_prio_current in ed_prio_opts else 1, key="ed_t_prio")
        with e4_b:
            ed_stat_opts = ["🟢 Completed", "🔵 In Progress", "⚪ Not Started", "🟠 At Risk", "🔴 Overdue"]
            ed_stat_current = next((x for x in ed_stat_opts if x.split(" ", 1)[-1] == str(task_rec.get("status") or "Not Started")), "⚪ Not Started")
            ed_stat = st.selectbox("Status (automated)", ed_stat_opts, index=ed_stat_opts.index(ed_stat_current), key="ed_t_stat", disabled=True, help="Calculated from completion and the task end date.")
        with e4_c:
            ed_type_opts=["💼 Task", "◆ Milestone", "Subtask"]
            ed_type_current = "◆ Milestone" if str(task_rec.get("task_type") or "Task")=="Milestone" else ("Subtask" if str(task_rec.get("task_type") or "Task")=="Subtask" else "💼 Task")
            ed_task_type = st.selectbox("Task Type *", ed_type_opts, index=ed_type_opts.index(ed_type_current), key="ed_t_type")
        with e4_d:
            st.selectbox("Parent Task", ["None"] + [f"{int(r.id)} · {r.name}" for _, r in T.iterrows() if int(r.id) != int(task_rec.get("id", -1))], key="ed_t_parent")

        e5_a, e5_b, e5_c = st.columns([0.9, 0.9, 1.4])
        with e5_a:
            st.markdown("<div style='font-size:12.5px;font-weight:600;color:#0F172A;margin-bottom:6px'>Milestone</div>", unsafe_allow_html=True)
            st.toggle("◆ Mark as Milestone", value=str(task_rec.get("task_type") or "Task")=="Milestone", key="ed_t_ms")
        with e5_b:
            st.markdown("<div style='font-size:12.5px;margin-bottom:6px'>&nbsp;</div>", unsafe_allow_html=True)
            st.selectbox("Milestone Type", ["Not a milestone", "Project Kick-off", "Design Sign-off", "Go-Live"], key="ed_t_ms_sel", label_visibility="collapsed")
        with e5_c:
            _existing_tags=[x.strip() for x in str(task_rec.get("tags") or "").split(",") if x.strip()]
            ed_tags = st.multiselect("🏷 Tags", ["Design", "UI/UX", "Planning", "Development", "Testing", "Launch", "Research", "Strategy"], default=_existing_tags, key="ed_t_tags")

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
            for _, child in ed_children.iterrows():
                child_key=f"ed_child_done_{int(child['id'])}"
                st.checkbox(str(child["name"]), value=str(child.get("status") or "") == "Completed", key=child_key)
            if st.session_state.get("ed_sub_items"):
                st.caption("New subtasks to add: " + ", ".join(st.session_state["ed_sub_items"]))
            if ed_children.empty and not st.session_state.get("ed_sub_items"):
                st.caption("No subtasks yet. Add a subtask above to enable automatic task progress.")
        with e6_r, st.container(border=True):
            st.markdown("<div style='font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px'>📎 &nbsp;Attachments</div><div style='font-size:12px;color:#64748B'>No attachments are stored for this task.</div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("<div style='font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px'>📄 &nbsp;Additional Information <span style='font-weight:400;color:#64748B'>(Optional)</span></div>", unsafe_allow_html=True)
            ai1, ai2, ai3, ai4 = st.columns([0.8, 0.8, 1.2, 1.4])
            with ai1:
                ed_est = st.number_input("Estimated Effort (Hours)", min_value=0.0, value=float(task_rec.get("estimated_effort") or 0), step=1.0, key="ed_est_h")
            with ai2:
                ed_act = st.number_input("Actual Effort (Hours)", min_value=0.0, value=float(task_rec.get("actual_effort") or 0), step=1.0, key="ed_act_h")
            with ai3:
                _dep_opts=["None"]+[f"{int(row.id)} · {row.name}" for _,row in T.iterrows() if int(row.id)!=ed_task_id]
                _dep_current=str(task_rec.get("dependency") or "None")
                if _dep_current not in _dep_opts: _dep_opts.append(_dep_current)
                ed_dep = st.selectbox("Dependencies", _dep_opts, index=_dep_opts.index(_dep_current), key="ed_dep_s")
            with ai4:
                ed_notes = st.text_area("Progress Notes", value=str(task_rec.get("notes") or ""), height=68, key="ed_p_notes")

        f_meta, f_can, f_sav = st.columns([3.3, 0.85, 1.15])
        with f_meta:
            st.markdown(
                clean_html(
                    f"""
                    <div style="display:flex;gap:24px;align-items:center;font-size:11.5px;color:#475569;padding-top:4px">
                        <div><span style="color:#64748B">Assigned to</span><br>{avatar_html(tass)}</div>
                        <div><span style="color:#64748B">Schedule</span><br><span style="color:#64748B">{html.escape(tstart)} – {html.escape(tend)}</span></div>
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
                if ed_end < ed_start:
                    st.error("End date cannot be earlier than start date.")
                else:
                    raw_status = str(ed_stat).split(" ", 1)[-1]
                    raw_priority = str(ed_prio).split(" ", 1)[-1]
                    selected_project_id = pmap.get(ed_proj, task_rec.get("project_id"))
                    parent_id = None
                    parent_choice = st.session_state.get("ed_t_parent", "None")
                    try:
                        if parent_choice != "None": parent_id = int(parent_choice.split(" ", 1)[0].split(".")[0])
                    except Exception: parent_id = None
                    ed_type_value = "Milestone" if st.session_state.get("ed_t_ms", False) or str(ed_task_type).endswith("Milestone") else ("Subtask" if ed_task_type == "Subtask" else "Task")
                    ed_dep_value = "" if ed_dep == "None" else ed_dep
                    original_notes = str(task_rec.get("notes") or "")
                    if ed_desc.strip() != original_notes and ed_notes.strip() != original_notes:
                        combined_notes = f"{ed_desc.strip()}\nProgress Notes: {ed_notes.strip()}".strip()
                    elif ed_desc.strip() != original_notes:
                        combined_notes = ed_desc.strip()
                    else:
                        combined_notes = ed_notes.strip()
                    sql("UPDATE tasks SET name=?,project_id=?,start_date=?,due_date=?,assignee=?,priority=?,status=?,notes=?,parent_task_id=?,dependency=?,task_type=?,tags=?,estimated_effort=?,actual_effort=? WHERE id=?",
                        (ed_name.strip(), selected_project_id, ed_start.isoformat(), ed_end.isoformat(), str(ed_ass).split(" (", 1)[0], raw_priority, raw_status, combined_notes, parent_id, ed_dep_value, ed_type_value, ",".join(ed_tags), float(ed_est), float(ed_act), ed_task_id))
                    for _, child in ed_children.iterrows():
                        child_key=f"ed_child_done_{int(child['id'])}"
                        child_done=bool(st.session_state.get(child_key, str(child.get("status") or "") == "Completed"))
                        sql("UPDATE tasks SET status=? WHERE id=?", ("Completed" if child_done else "Not Started", int(child["id"])))
                    for new_subtask in st.session_state.get("ed_sub_items", []):
                        sql("INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                            (new_subtask, selected_project_id, str(task_rec.get("team") or "Portfolio"), "Not Started", raw_priority, ed_start.isoformat(), ed_end.isoformat(), str(ed_ass).split(" (", 1)[0], 0, "", "Checklist subtask", ed_task_id))
                    existing_milestone = df("SELECT id FROM milestones WHERE project_id=? AND name=?", (selected_project_id, ed_name.strip()))
                    if ed_type_value == "Milestone":
                        if existing_milestone.empty:
                            sql("INSERT INTO milestones(name,project_id,team,status,start_date,due_date,completion_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?)",
                                (ed_name.strip(), selected_project_id, str(task_rec.get("team") or "Portfolio"), raw_status, ed_start.isoformat(), ed_end.isoformat(), date.today().isoformat() if raw_status == "Completed" else "", ed_dep_value, ed_notes))
                        else:
                            sql("UPDATE milestones SET status=?,start_date=?,due_date=?,dependency=?,notes=? WHERE id=?",(raw_status,ed_start.isoformat(),ed_end.isoformat(),ed_dep_value,ed_notes,int(existing_milestone.iloc[0]["id"])))
                    elif not existing_milestone.empty:
                        sql("DELETE FROM milestones WHERE id=?",(int(existing_milestone.iloc[0]["id"]),))
                    st.session_state["ed_sub_items"] = []
                    log_event("Updated", "Task", ed_task_id, ed_name)
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
        tl_date_range = st.date_input("Date Range", (date.today(), date.today() + timedelta(days=60)), key="tl_d_top", label_visibility="collapsed")
    with th_b:
        if st.button("＋ Add Task", type="primary", use_container_width=True, key="tl_btn_add_top"):
            if not projects_list:
                st.warning("Import the portfolio workbook in Settings or create a project before adding tasks.")
            else:
                for _key in list(st.session_state.keys()):
                    if _key.startswith(("add_t_", "add_t_sub_")):
                        st.session_state.pop(_key, None)
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
        tl_ass_f = st.selectbox("Assignee", ["All Assignees"] + sorted(R.name.dropna().astype(str).tolist()), key="tl_f_ass", label_visibility="collapsed")
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
    # --- KPI CARDS (calculated from live task data) ---
    task_counts = T["status"].fillna("Not Started").value_counts().to_dict() if not T.empty else {}
    task_total = len(T)
    kpi_specs = [
        ("▤", "Total Tasks", task_total, "#1473E6"),
        ("✓", "Completed", int(task_counts.get("Completed", 0)), "#00875A"),
        ("❂", "In Progress", int(task_counts.get("In Progress", 0)), "#1473E6"),
        ("●", "Not Started", int(task_counts.get("Not Started", 0)), "#64748B"),
        ("!", "At Risk", int(task_counts.get("At Risk", 0)), "#F28C00"),
        ("!", "Overdue", int(task_counts.get("Overdue", 0)), "#EF4444"),
    ]
    kpi_cols = st.columns(6)
    for col, (ico, label, value, color) in zip(kpi_cols, kpi_specs):
        with col:
            pct = round(value * 100 / task_total) if task_total else 0
            st.markdown(clean_html(f"<div class='vk-kpi-card' style='min-height:84px;padding:12px 14px'><div style='width:40px;height:40px;min-width:40px;border-radius:50%;background:{color};color:white;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800'>{ico}</div><div><div style='font-size:11.5px;font-weight:600;color:#475569'>{label}</div><div style='font-size:21px;font-weight:800;color:#0F172A;line-height:1.2'>{value} <span style='font-size:12px;font-weight:600'>({pct}%)</span></div></div></div>"), unsafe_allow_html=True)

    st.write("")
    # --- GROUPED WBS TASK TABLE (live DB data; project progress is based on completed leaf tasks) ---
    task_groups = []
    if not T.empty:
        children_ids = set(pd.to_numeric(T.get("parent_task_id", pd.Series(dtype=float)), errors="coerce").dropna().astype(int).tolist())
        for _, pr in P.iterrows():
            pr_tasks = T[T["project_id"] == pr["id"]].copy()
            if isinstance(tl_date_range, tuple) and len(tl_date_range) == 2 and not pr_tasks.empty:
                _due_filter = pd.to_datetime(pr_tasks["due_date"], errors="coerce")
                pr_tasks = pr_tasks[_due_filter.isna() | ((_due_filter.dt.date >= tl_date_range[0]) & (_due_filter.dt.date <= tl_date_range[1]))]
            if pr_tasks.empty:
                continue
            pr_tasks["_due"] = pd.to_datetime(pr_tasks["due_date"], errors="coerce")
            pr_tasks = pr_tasks.sort_values(["_due", "id"], na_position="last")
            try: pstart = date.fromisoformat(str(pr.get("start_date") or "")[:10])
            except Exception: pstart = None
            try: pend = date.fromisoformat(str(pr.get("end_date") or "")[:10])
            except Exception: pend = None
            p_days = (pend - pstart).days + 1 if pstart and pend else 0
            group_tasks = []
            for _, tr in pr_tasks.iterrows():
                sid = tr.get("source_id")
                wbs = str(sid) if pd.notna(sid) and str(sid).strip() else str(int(tr["id"]))
                sd = fmt_date(tr.get("start_date")); ed = fmt_date(tr.get("due_date"))
                try:
                    dstart = date.fromisoformat(str(tr.get("start_date"))[:10])
                    dend = date.fromisoformat(str(tr.get("due_date"))[:10])
                    dur = f"{max(1,(dend-dstart).days+1)} days"
                except Exception: dur = "—"
                is_parent = int(tr["id"]) in children_ids
                status_value = str(tr.get("status") or "Not Started")
                priority_value = str(tr.get("priority") or "Medium")
                assignee_value = str(tr.get("assignee") or "Unassigned")
                group_tasks.append({"id": int(tr["id"]), "wbs": wbs, "ico": "doc" if is_parent else "play", "ic_col": GREEN if status_value == "Completed" else BLUE,
                    "name": str(tr.get("name") or "Untitled Task"), "project": str(pr["name"]), "start_date": sd, "end_date": ed, "duration": dur,
                    "status": status_value, "priority": priority_value, "assignee": assignee_value, "pct": int(tr.get("progress") or 0),
                    "bar_col": GREEN if status_value == "Completed" else (RED if status_value == "Overdue" else BLUE), "tag": str(tr.get("team") or "Portfolio"),
                    "tag_bg": "#DBEAFE", "tag_fg": "#1D4ED8", "has_children": is_parent, "project_id": int(pr["id"]), "start_iso": tr.get("start_date"), "due_iso": tr.get("due_date"), "parent_task_id": tr.get("parent_task_id"), "notes": tr.get("notes"), "team": tr.get("team"), "tags": tr.get("tags"), "task_type": tr.get("task_type"), "dependency": tr.get("dependency"), "estimated_effort": tr.get("estimated_effort"), "actual_effort": tr.get("actual_effort")})
            project_pct = int(pr.get("progress") or 0)
            task_groups.append({"group_title": f"{pr['name']} ({len(pr_tasks)})", "project": str(pr["name"]), "icon_bg": str(pr.get("bar_color") or GREEN),
                "start": fmt_date(pr.get("start_date")), "end": fmt_date(pr.get("end_date")), "dur": f"{p_days} days" if p_days else "—", "pct": project_pct,
                "bar_col": str(pr.get("bar_color") or GREEN), "row_bg": "#F0FDF4", "tasks": group_tasks})

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
                    is_done = t_item["status"] == "Completed"
                    mark_done = st.checkbox("", value=is_done, key=f"tl_done_{t_item['id']}", disabled=bool(t_item.get("has_children")), help="Completion is derived from child tasks" if t_item.get("has_children") else "Mark task complete")
                    if not t_item.get("has_children") and mark_done != is_done:
                        sql("UPDATE tasks SET status=? WHERE id=?", ("Completed" if mark_done else "Not Started", t_item["id"]))
                        st.rerun()
                    st.caption(wbs_id)
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
                    if t_item.get("has_children"):
                        st.markdown(clean_html(f"<div style='display:flex;align-items:center;gap:8px;padding-top:6px'><span style='font-size:12px;font-weight:600;color:#1E293B;min-width:30px'>{pct_v}%</span><div style='flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden'><div style='width:{pct_v}%;height:100%;background:{t_item['bar_col']};border-radius:99px'></div></div></div>"), unsafe_allow_html=True)
                    else:
                        st.markdown("<div style='padding-top:6px;color:#94A3B8'>—</div>", unsafe_allow_html=True)
                with rc[10]:
                    st.markdown(
                        f"<div style='padding-top:4px'><span class='vk-tag' style='background:{t_item['tag_bg']};color:{t_item['tag_fg']}'>{html.escape(t_item['tag'])}</span></div>",
                        unsafe_allow_html=True,
                    )
                with rc[11]:
                    b_v, b_e = st.columns(2)
                    with b_v:
                        if st.button("👁", key=f"tl_v_{t_item['id']}", help="View Task"):
                            st.session_state.task_modal = ("view", t_item)
                            st.rerun()
                    with b_e:
                        if st.button("❐", key=f"tl_e_{t_item['id']}", help="Edit Task"):
                            st.session_state.task_modal = ("edit", t_item)
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
        resource_name=str(res_rec.get("name") or "")
        if "rc_calendar_month" not in st.session_state: st.session_state["rc_calendar_month"]=date.today().replace(day=1)
        month=st.session_state["rc_calendar_month"]
        h1,h2,h3,h4=st.columns([2,0.7,0.7,1])
        with h1: st.markdown(f"<div style='font-size:20px;font-weight:800'>{html.escape(resource_name)} — Calendar</div><div style='font-size:12px;color:#64748B'>{month.strftime('%B %Y')}</div>",unsafe_allow_html=True)
        with h2:
            if st.button("‹ Prev",key="rc_month_prev",use_container_width=True):
                st.session_state["rc_calendar_month"]=(month.replace(day=1)-timedelta(days=1)).replace(day=1); st.rerun()
        with h3:
            if st.button("Next ›",key="rc_month_next",use_container_width=True):
                st.session_state["rc_calendar_month"]=(month.replace(day=28)+timedelta(days=4)).replace(day=1); st.rerun()
        with h4:
            if st.button("Today",key="rc_today_btn",use_container_width=True): st.session_state["rc_calendar_month"]=date.today().replace(day=1); st.rerun()
        month_end=(month.replace(day=28)+timedelta(days=4)).replace(day=1)-timedelta(days=1)
        resource_tasks=T[T["assignee"].astype(str)==resource_name].copy() if not T.empty else pd.DataFrame()
        if not resource_tasks.empty:
            due=pd.to_datetime(resource_tasks["due_date"],errors="coerce")
            resource_tasks=resource_tasks[due.notna() & (due.dt.date>=month) & (due.dt.date<=month_end)]
        resource_events=df("SELECT * FROM calendar_events WHERE date(end_date)>=date(?) AND date(start_date)<=date(?) ORDER BY start_date,start_time",(month.isoformat(),month_end.isoformat()))
        if not resource_events.empty:
            resource_events=resource_events[resource_events["resources"].fillna("[]").apply(lambda v: resource_name in (json.loads(v) if str(v).strip().startswith("[") else []))]
        tab1,tab2=st.tabs([f"Tasks ({len(resource_tasks)})",f"Events ({len(resource_events)})"])
        with tab1:
            if resource_tasks.empty: st.info("No tasks due for this resource in the selected month.")
            else:
                cols=[c for c in ["name","project","start_date","due_date","status","priority","progress"] if c in resource_tasks.columns]
                st.dataframe(resource_tasks[cols],hide_index=True,use_container_width=True,height=300)
        with tab2:
            if resource_events.empty: st.info("No calendar events for this resource in the selected month.")
            else: st.dataframe(resource_events[[c for c in ["title","event_type","start_date","end_date","start_time","end_time","description"] if c in resource_events.columns]],hide_index=True,use_container_width=True,height=300)
        b1,b2=st.columns(2)
        with b1:
            if st.button("＋ Add Event",key="rc_add_ev_btn",type="primary",use_container_width=True):
                st.session_state.page="Calendar"; st.session_state.cal_modal="add_event"; st.session_state["ae_res"]=[resource_name]; st.rerun()
        with b2:
            if st.button("Close",key="rc_close_btn",use_container_width=True): st.session_state.pop("res_modal",None); st.rerun()

    @st.dialog("Project Assignments", width="large")
    def project_assignments_modal(res_rec: dict) -> None:
        rname = res_rec.get("name", "Unassigned")
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

        assigned_tasks = T[T["assignee"].astype(str) == str(res_rec.get("name", ""))].copy() if not T.empty else pd.DataFrame()
        pa_h1, pa_h2 = st.columns([3.5, 1.1])
        with pa_h1:
            st.markdown(f"<div style='font-size:14.5px;font-weight:700;color:#0F172A;padding-top:6px'>💼 &nbsp;Current Project Assignments ({assigned_tasks['project_id'].nunique() if not assigned_tasks.empty else 0})</div>", unsafe_allow_html=True)
        with pa_h2:
            if st.button("＋ Assign to Project", type="primary", key="pa_assign_btn", use_container_width=True):
                st.session_state.res_modal = ("assign", res_rec)
                st.rerun()
        if assigned_tasks.empty:
            st.info("No project assignments yet. Assign this resource to a project or allocate tasks from the task list.")
        else:
            assignment_rows=[]
            for project_name, group in assigned_tasks.groupby("project", dropna=False):
                p_row=P[P["name"]==project_name]
                completed=int((group["status"].astype(str)=="Completed").sum())
                total=len(group)
                pct=round(completed*100/total) if total else 0
                start_dates=pd.to_datetime(group["start_date"],errors="coerce").dropna()
                end_dates=pd.to_datetime(group["due_date"],errors="coerce").dropna()
                assignment_rows.append({"Project":project_name,"Assigned Tasks":total,"Completed":completed,"Task Progress":pct,
                    "Start Date":start_dates.min().strftime("%b %d, %Y").replace(" 0"," ") if not start_dates.empty else "—",
                    "End Date":end_dates.max().strftime("%b %d, %Y").replace(" 0"," ") if not end_dates.empty else "—",
                    "Status":str(p_row.iloc[0]["status"]) if not p_row.empty else str(group.iloc[0]["status"])})
            assign_df=pd.DataFrame(assignment_rows)
            st.dataframe(assign_df,hide_index=True,use_container_width=True,height=170,
                column_config={"Task Progress":st.column_config.ProgressColumn("Task Progress",min_value=0,max_value=100,format="%d%%")})
        with st.container(border=True):
            st.markdown("<div style='font-size:14.5px;font-weight:700;color:#0F172A;margin-bottom:6px'>🕒 &nbsp;Workload Overview</div>", unsafe_allow_html=True)
            wo_l, wo_m, wo_r = st.columns([0.9, 1.15, 1.85])
            if assigned_tasks.empty:
                with wo_l: st.info("No allocation data")
                with wo_m: st.caption("Allocation will appear as assignments are added.")
                with wo_r: st.caption("Weekly task schedule is derived from task dates.")
            else:
                project_counts=assigned_tasks.groupby("project").size().reset_index(name="Tasks")
                with wo_l:
                    fig_pa=px.pie(project_counts,names="project",values="Tasks",hole=0.68)
                    fig_pa.update_layout(height=165,margin=dict(l=4,r=4,t=4,b=4),showlegend=False,paper_bgcolor="rgba(0,0,0,0)",annotations=[dict(text=f"<b>{len(assigned_tasks)}</b><br><span style='font-size:10px;color:#64748B'>Assigned Tasks</span>",x=0.5,y=0.5,showarrow=False)])
                    st.plotly_chart(fig_pa,use_container_width=True,config={"displayModeBar":False})
                with wo_m:
                    st.dataframe(project_counts,hide_index=True,use_container_width=True,height=165)
                with wo_r:
                    wk=assigned_tasks.copy(); wk["due_dt"]=pd.to_datetime(wk["due_date"],errors="coerce")
                    wk=wk[wk["due_dt"].notna()].copy()
                    if wk.empty: st.info("No dated tasks for this resource.")
                    else:
                        wk["Weekday"]=wk["due_dt"].dt.day_name().str[:3]
                        wk_df=wk.groupby(["Weekday","project"]).size().reset_index(name="Tasks")
                        fig_wk=px.bar(wk_df,x="Weekday",y="Tasks",color="project",barmode="stack")
                        fig_wk.update_layout(height=165,margin=dict(l=0,r=0,t=4,b=0),xaxis_title=None,yaxis_title=None,paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_wk,use_container_width=True,config={"displayModeBar":False})

        assigned_tasks = T[T["assignee"].astype(str) == str(res_rec.get("name", ""))].copy() if not T.empty else pd.DataFrame()
        pa_tabs = st.tabs([f"📅 Assigned Tasks ({len(assigned_tasks)})", "☆ Milestones", "🕒 Time Allocation"])
        with pa_tabs[0]:
            if assigned_tasks.empty:
                st.info("No tasks are assigned to this resource yet.")
            else:
                view_cols = [c for c in ["name", "project", "start_date", "due_date", "status", "priority", "progress"] if c in assigned_tasks.columns]
                st.dataframe(assigned_tasks[view_cols], hide_index=True, use_container_width=True, height=260)
        with pa_tabs[1]:
            res_milestones = M[M.get("team", pd.Series(index=M.index, dtype=str)).astype(str) == str(res_rec.get("team", ""))] if not M.empty else pd.DataFrame()
            if res_milestones.empty: st.info("No milestones are assigned to this resource's team.")
            else: st.dataframe(res_milestones[[c for c in ["name", "project", "status", "due_date"] if c in res_milestones.columns]], hide_index=True, use_container_width=True)
        with pa_tabs[2]:
            total_assigned = len(assigned_tasks)
            completed_assigned = int((assigned_tasks["status"].astype(str) == "Completed").sum()) if total_assigned else 0
            completion_pct = round(completed_assigned * 100 / total_assigned) if total_assigned else 0
            allocation = int(res_rec.get("capacity") or 0)
            st.metric("Task Completion", f"{completion_pct}%", f"{completed_assigned} of {total_assigned} tasks completed")
            st.progress(completion_pct / 100 if total_assigned else 0.0)
            st.metric("Capacity / Allocation", f"{allocation}%")

            if st.button("Close", key="pa_close_btn", use_container_width=True):
                st.session_state.pop("res_modal", None)
                st.rerun()

    @st.dialog("Resource Action")
    def resource_action_submodal(act_mode: str, res_rec: dict) -> None:
        rname = res_rec.get("name", "Unassigned")
        rrole = res_rec.get("role", "UX Designer")
        remail = res_rec.get("email", "james.miller@hpe.com")
        rini = res_rec.get("ini", "JM")
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
            resource_id=int(res_rec.get("id") or 0)
            existing_assignments=df("SELECT ra.*,p.name AS project_name FROM resource_assignments ra LEFT JOIN projects p ON p.id=ra.project_id WHERE ra.resource_id=?",(resource_id,)) if resource_id else pd.DataFrame()
            current_projects=existing_assignments["project_name"].dropna().astype(str).tolist() if not existing_assignments.empty else []
            selected_projects=st.multiselect("Select Project(s) *",projects_list,default=[x for x in current_projects if x in projects_list],key="sub_as_projects")
            allocation_value=st.number_input("Allocation (%)",min_value=0,max_value=200,value=int(ralloc or 0),step=10,key="sub_as_pct")
            assign_start=st.date_input("Assignment Start Date",date.today(),key="sub_as_start")
            assign_end=st.date_input("Assignment End Date",date.today()+timedelta(days=60),key="sub_as_end")
            c1,c2=st.columns(2)
            with c1:
                if st.button("Cancel",key="sub_as_can",use_container_width=True): st.session_state.pop("res_modal",None); st.rerun()
            with c2:
                if st.button("Assign",type="primary",key="sub_as_ok",use_container_width=True):
                    if not selected_projects:
                        st.error("Select at least one project.")
                    elif assign_end < assign_start:
                        st.error("Assignment end date cannot be earlier than start date.")
                    elif not resource_id:
                        st.error("This resource record has no valid ID.")
                    else:
                        with conn() as c:
                            selected_ids=[int(pmap[name]) for name in selected_projects if name in pmap]
                            for project_id in selected_ids:
                                project_end=df("SELECT end_date FROM projects WHERE id=?",(project_id,))
                                end_value=assign_end.isoformat()
                                c.execute("INSERT INTO resource_assignments(resource_id,project_id,allocation,start_date,end_date) VALUES(?,?,?,?,?) ON CONFLICT(resource_id,project_id) DO UPDATE SET allocation=excluded.allocation,start_date=excluded.start_date,end_date=excluded.end_date",(resource_id,project_id,int(allocation_value),assign_start.isoformat(),end_value))
                            existing_ids=[int(x) for x in existing_assignments["project_id"].tolist()] if not existing_assignments.empty else []
                            for project_id in set(existing_ids)-set(selected_ids): c.execute("DELETE FROM resource_assignments WHERE resource_id=? AND project_id=?",(resource_id,project_id))
                            c.commit()
                        _persist_snapshot(); log_event("Assigned", "Resource", resource_id, ", ".join(selected_projects))
                        st.session_state.pop("res_modal",None); st.rerun()

        elif act_mode == "skills":
            st.text_input("Select Skills", placeholder="🔍 Search skills...", key="sub_sk_q")
            sk_c1, sk_c2 = st.columns(2)
            with sk_c1:
                st.checkbox("UI/UX", value="UI/UX" in res_rec.get("skills", []), key="sk_1")
                st.checkbox("React", value="React" in res_rec.get("skills", []), key="sk_2")
                st.checkbox("API", value="API" in res_rec.get("skills", []), key="sk_3")
                st.checkbox("Testing", value="Testing" in res_rec.get("skills", []), key="sk_4")
            with sk_c2:
                st.checkbox("Figma", value="Figma" in res_rec.get("skills", []), key="sk_5")
                st.checkbox("JavaScript", value="JavaScript" in res_rec.get("skills", []), key="sk_6")
                st.checkbox("Data Analysis", value="Data Analysis" in res_rec.get("skills", []), key="sk_7")
                st.checkbox("Automation", value="Automation" in res_rec.get("skills", []), key="sk_8")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_sk_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Save Skills", type="primary", key="sub_sk_ok", use_container_width=True):
                    selected_skills=[]
                    for _skill_key,_skill_label in [("sk_1","UI/UX"),("sk_2","React"),("sk_3","API"),("sk_4","Testing"),("sk_5","Figma"),("sk_6","JavaScript"),("sk_7","Data Analysis"),("sk_8","Automation")]:
                        if st.session_state.get(_skill_key,False): selected_skills.append(_skill_label)
                    sql("UPDATE resources SET skills=? WHERE id=?",(",".join(selected_skills),int(res_rec.get("id") or 0)))
                    log_event("Updated Skills", "Resource", int(res_rec.get("id") or 0), ", ".join(selected_skills))
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "availability":
            av_start = st.date_input("Availability Start Date *", date.today(), key="sub_av_st")
            av_has_end = st.checkbox("Set availability end date", value=False, key="sub_av_has_end")
            av_end = st.date_input("Availability End Date", date.today()+timedelta(days=30), key="sub_av_en") if av_has_end else None
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
                    if av_end and av_end < av_start:
                        st.error("Availability end date cannot be earlier than start date.")
                    else:
                        sql("UPDATE resources SET capacity=? WHERE name=?", (int(new_cap), rname))
                        sql("INSERT INTO calendar_events(title,event_type,resources,start_date,end_date,allocation,description) VALUES(?,?,?,?,?,?,?)",(f"{rname} — Availability","Availability",json.dumps([rname]),av_start.isoformat(),(av_end or av_start).isoformat(),int(new_cap),f"Working hours: {int(st.session_state.get('sub_av_hr',8))} per day"))
                        log_event("Updated Availability", "Resource", int(res_rec.get("id") or 0), rname)
                        st.session_state.pop("res_modal", None)
                        st.rerun()

        elif act_mode == "leave":
            lv1, lv2 = st.columns(2)
            with lv1:
                lv_start = st.date_input("Leave Start Date *", date.today(), key="sub_lv_st")
            with lv2:
                lv_end = st.date_input("Leave End Date *", date.today() + timedelta(days=9), key="sub_lv_en")
            st.selectbox("Reason (Optional)", ["Vacation", "Sick Leave", "Personal Leave", "Training"], key="sub_lv_rs")
            st.text_input("Notes (Optional)", placeholder="Enter notes...", key="sub_lv_nt")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_lv_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Set on Leave", type="primary", key="sub_lv_ok", use_container_width=True):
                    if lv_end < lv_start:
                        st.error("Leave end date cannot be earlier than start date.")
                    else:
                        sql("UPDATE resources SET status='On Leave', capacity=0 WHERE name=?", (rname,))
                        sql("INSERT INTO calendar_events(title,event_type,resources,start_date,end_date,description) VALUES(?,?,?,?,?,?)",(f"{rname} — Leave","Leave",json.dumps([rname]),lv_start.isoformat(),lv_end.isoformat(),f"{st.session_state.get('sub_lv_rs','Leave')} · {st.session_state.get('sub_lv_nt','')}"))
                        log_event("Set on Leave", "Resource", int(res_rec.get("id") or 0), rname)
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
        res_date_range = st.date_input("Date Range", (date.today(), date.today() + timedelta(days=60)), key="res_d_top", label_visibility="collapsed")
    with rh_b:
        if st.button("＋ Add Resource", type="primary", use_container_width=True, key="res_btn_add_top"):
            st.session_state.res_modal = ("add", None)
            st.rerun()

    st.write("")
    # --- RESOURCE KPI CARDS (live records) ---
    _resource_states=[]
    for _,_res in R.iterrows():
        _assigned=T[T["assignee"].astype(str)==str(_res.get("name",""))] if not T.empty else pd.DataFrame()
        _ar=df("SELECT SUM(allocation) AS alloc FROM resource_assignments WHERE resource_id=?",(int(_res.get("id") or 0),))
        _capacity=int(_res.get("capacity") or 0)
        _alloc=int(_ar.iloc[0]["alloc"] or 0) if not _ar.empty and pd.notna(_ar.iloc[0]["alloc"]) else (min(150,int((~_assigned["status"].astype(str).eq("Completed")).sum())*20) if not _assigned.empty else 0)
        if str(_res.get("status") or "").lower()=="on leave": _state="On Leave"
        elif _alloc>_capacity: _state="Overallocated"
        elif _alloc>0: _state="Allocated"
        else: _state="Available"
        _resource_states.append(_state)
    _state_counts=pd.Series(_resource_states).value_counts().to_dict() if _resource_states else {}
    _total_res=len(R)
    rk1,rk2,rk3,rk4,rk5=st.columns(5)
    res_kpis=[(rk1,"#1473E6","👥","Total Resources",_total_res),(rk2,"#00875A","✓","Available",int(_state_counts.get("Available",0))),(rk3,"#1473E6","🕒","Allocated",int(_state_counts.get("Allocated",0))),(rk4,"#F28C00","!","Overallocated",int(_state_counts.get("Overallocated",0))),(rk5,"#64748B","👤","On Leave",int(_state_counts.get("On Leave",0)))]
    for col,ibg,ico,lbl,val in res_kpis:
        with col:
            pct_s=f"({round(val*100/_total_res) if _total_res else 0}%)" if lbl!="Total Resources" else ""
            st.markdown(clean_html(f"<div class='vk-kpi-card' style='min-height:86px;padding:14px 16px'><div style='width:44px;height:44px;min-width:44px;border-radius:50%;background:{ibg};color:white;display:flex;align-items:center;justify-content:center;font-size:19px;font-weight:800'>{ico}</div><div><div style='font-size:12px;font-weight:600;color:#334155'>{lbl}</div><div style='font-size:23px;font-weight:800;color:#0F172A;line-height:1.2'>{val} <span style='font-size:18px;font-weight:800'>{pct_s}</span></div></div></div>"),unsafe_allow_html=True)

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
        _months=[(date.today().replace(day=1)+timedelta(days=31*i)).replace(day=1) for i in range(3)]
        _month_names=[d.strftime("%b %Y") for d in _months]
        _resource_count=max(1,len(R))
        _leave_count=int((R.get("status",pd.Series(dtype=str)).astype(str)=="On Leave").sum()) if not R.empty else 0
        _available_count=int((R.get("status",pd.Series(dtype=str)).astype(str)=="Available").sum()) if not R.empty else 0
        ra_df=pd.DataFrame({"Month":_month_names,"Billable":[max(0,len(T)*0.6/_resource_count)]*3,"Non-Billable":[max(0,len(T)*0.25/_resource_count)]*3,"Leave":[_leave_count]*3})
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
            _capacity_total=max(1,int(pd.to_numeric(R.get("capacity",pd.Series(dtype=float)),errors="coerce").fillna(0).sum()))
            _manual_total=df("SELECT SUM(allocation) AS total FROM resource_assignments")
            _alloc_total=int(_manual_total.iloc[0]["total"] or 0) if not _manual_total.empty and pd.notna(_manual_total.iloc[0]["total"]) else int(sum(20 for _ in range(len(T[ T["status"].astype(str)!="Completed"]))))
            _util_pct=min(100,round(_alloc_total*100/_capacity_total))
            _nonbill_pct=round(_util_pct*0.3)
            _available_pct=max(0,100-_util_pct-_nonbill_pct)
            fig_ut = go.Figure(
                go.Pie(
                    values=[_util_pct-_nonbill_pct, _nonbill_pct, _available_pct],
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
                        text=f"<b style='font-size:20px;color:#0F172A'>{_util_pct}%</b><br><span style='font-size:10.5px;color:#475569'>Team Utilization</span>",
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
                    f"<div style='font-size:12.5px;line-height:2.4;padding-top:32px'><div style='display:flex;justify-content:space-between'><span><span style='color:{GREEN}'>●</span> &nbsp;Billable</span><b>{_util_pct-_nonbill_pct}%</b></div><div style='display:flex;justify-content:space-between'><span><span style='color:{BLUE}'>●</span> &nbsp;Non-Billable</span><b>{_nonbill_pct}%</b></div><div style='display:flex;justify-content:space-between'><span><span style='color:#94A3B8'>●</span> &nbsp;Available</span><b>{_available_pct}%</b></div></div>"
                ),
                unsafe_allow_html=True,
            )

    with rc_3, st.container(border=True):
        st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;margin-bottom:12px'>Resource Status</div>", unsafe_allow_html=True)
        rs_bars = [("Available",int(_state_counts.get("Available",0)),round(_state_counts.get("Available",0)*100/max(1,_total_res)),"#6EE7B7"),("Allocated",int(_state_counts.get("Allocated",0)),round(_state_counts.get("Allocated",0)*100/max(1,_total_res)),"#1473E6"),("Overallocated",int(_state_counts.get("Overallocated",0)),round(_state_counts.get("Overallocated",0)*100/max(1,_total_res)),"#F87171"),("On Leave",int(_state_counts.get("On Leave",0)),round(_state_counts.get("On Leave",0)*100/max(1,_total_res)),"#E2E8F0")]
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
        res_skill_f = st.selectbox("Skill / Team", ["All Skills"] + sorted(R.team.dropna().astype(str).unique().tolist()), key="res_f_skill", label_visibility="collapsed")
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
    # --- RESOURCES DIRECTORY TABLE (live resource records and task-based allocation) ---
    res_directory=[]
    for _,res in R.iterrows():
        assigned=T[T["assignee"].astype(str)==str(res.get("name",""))].copy() if not T.empty else pd.DataFrame()
        if isinstance(res_date_range,tuple) and len(res_date_range)==2 and not assigned.empty:
            _adue=pd.to_datetime(assigned["due_date"],errors="coerce")
            assigned=assigned[_adue.isna()|((_adue.dt.date>=res_date_range[0])&(_adue.dt.date<=res_date_range[1]))]
        task_projects=sorted(assigned["project"].dropna().astype(str).unique().tolist()) if not assigned.empty and "project" in assigned else []
        assignment_rows=df("SELECT ra.allocation,p.name AS project_name FROM resource_assignments ra LEFT JOIN projects p ON p.id=ra.project_id WHERE ra.resource_id=?",(int(res.get("id") or 0),))
        manual_projects=assignment_rows["project_name"].dropna().astype(str).tolist() if not assignment_rows.empty else []
        project_names=sorted(set(task_projects+manual_projects))
        active_count=int((~assigned["status"].astype(str).eq("Completed")).sum()) if not assigned.empty else 0
        alloc=int(assignment_rows["allocation"].fillna(0).sum()) if not assignment_rows.empty else min(150,active_count*20)
        capacity=int(res.get("capacity") or 0)
        available=max(0,capacity-alloc)
        status=str(res.get("status") or "Available")
        if status.lower()=="on leave": status="On Leave"
        elif alloc>capacity: status="Overallocated"
        elif alloc>0: status="Allocated"
        else: status="Available"
        nm=str(res.get("name") or "Unassigned")
        initials="".join(part[0] for part in nm.split()[:2]).upper()
        team=str(res.get("team") or "")
        stored_skills=[x.strip() for x in str(res.get("skills") or "").split(",") if x.strip()]
        skill_list=stored_skills or ([team] if team else [])
        res_directory.append({"id":int(res.get("id") or 0),"ini":initials,"av_col":"#0284C7","name":nm,"email":str(res.get("email") or ""),
            "role":str(res.get("role") or "Team Member"),"skills":skill_list,"sk_cols":[("#DBEAFE","#1D4ED8") for _ in skill_list],"project":", ".join(project_names) if project_names else "Unassigned",
            "alloc":alloc,"avail":f"{available}%","status":status,"team":team,"capacity":capacity,"notes":str(res.get("notes") or "")})
    q_res = (res_top_q + " " + res_sub_q).strip().lower()
    filtered_res=[]
    for item in res_directory:
        if res_role_f!="All Roles" and item["role"]!=res_role_f: continue
        if res_skill_f!="All Skills" and res_skill_f not in item["skills"]: continue
        if res_stat_f!="All Statuses" and item["status"]!=res_stat_f: continue
        if res_proj_f!="All Projects" and res_proj_f not in item["project"].split(", "): continue
        if q_res and not (q_res in item["name"].lower() or q_res in item["role"].lower() or q_res in item["project"].lower() or q_res in item["team"].lower()): continue
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

        st.multiselect(
            "Resource(s) *", R.name.dropna().astype(str).tolist(),
            default=R.name.dropna().astype(str).tolist()[:1], key="ae_res",
        )

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
                if not ev_title.strip():
                    st.error("Enter an event title.")
                elif ev_end_d < ev_start_d:
                    st.error("End date cannot be earlier than start date.")
                else:
                    p_id = pmap.get(ev_proj) if ev_proj != "Select project" else None
                    task_name = st.session_state.get("ae_task", "Select task (optional)")
                    task_id = int(T.loc[T.name == task_name, "id"].iloc[0]) if task_name != "Select task (optional)" and not T.loc[T.name == task_name].empty else None
                    type_clean = ev_type.split(" ", 1)[-1]
                    sql("INSERT INTO calendar_events(title,event_type,project_id,task_id,resources,start_date,end_date,start_time,end_time,allocation,repeat_rule,description) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                        (ev_title.strip(), type_clean, p_id, task_id, json.dumps(st.session_state.get("ae_res", [])), ev_start_d.isoformat(), ev_end_d.isoformat(), st.session_state.get("ae_st", "09:00 AM"), st.session_state.get("ae_et", "05:00 PM"), int(float(st.session_state.get("ae_alloc", "0") or 0)), st.session_state.get("ae_rep", "Does not repeat"), ev_desc.strip()))
                    log_event("Created", "Calendar Event", None, ev_title.strip())
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
        cal_res_f = st.selectbox("Resource", ["All Resources"] + R.name.dropna().astype(str).tolist(), key="cal_f_res", label_visibility="collapsed")
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
    with ct_nav:
        if "cal_calendar_month" not in st.session_state:
            st.session_state["cal_calendar_month"] = date.today().replace(day=1)
        nav_l, nav_m, nav_r = st.columns([0.5, 2, 0.5])
        with nav_l:
            if st.button("‹", key="cal_month_prev", use_container_width=True):
                cur = st.session_state["cal_calendar_month"]
                st.session_state["cal_calendar_month"] = (cur.replace(day=1) - timedelta(days=1)).replace(day=1)
                st.session_state["cal_selected_date_input"] = st.session_state["cal_calendar_month"]
                st.session_state["cal_selected_date"] = st.session_state["cal_calendar_month"]
                st.rerun()
        with nav_m:
            st.markdown(f"<div style='text-align:center;font-weight:800;padding-top:7px'>{st.session_state['cal_calendar_month'].strftime('%B %Y')}</div>", unsafe_allow_html=True)
        with nav_r:
            if st.button("›", key="cal_month_next", use_container_width=True):
                cur = st.session_state["cal_calendar_month"]
                nxt = (cur.replace(day=28) + timedelta(days=4)).replace(day=1)
                st.session_state["cal_calendar_month"] = nxt
                st.session_state["cal_selected_date_input"] = nxt
                st.session_state["cal_selected_date"] = nxt
                st.rerun()
    with ct_tod:
        if st.button("Today", key="cal_tod_btn", use_container_width=True):
            st.session_state["cal_selected_date"] = date.today()
            st.session_state["cal_selected_date_input"] = date.today()
            st.session_state["cal_calendar_month"] = date.today().replace(day=1)
            st.rerun()
    with ct_q:
        cal_search_q = st.text_input("Search", placeholder="🔍 Search events, tasks, or team members...", key="cal_q_box", label_visibility="collapsed")
    with ct_mf:
        if st.button("▽ More Filters", key="cal_more_btn", use_container_width=True):
            st.session_state["cal_q_box"] = ""
            st.rerun()

    st.write("")
    # --- MAIN SPLIT: LEFT (Calendar Grid + 4 Bottom Cards) | RIGHT (Selected Date Inspector) ---
    main_l, main_r = st.columns([2.75, 1.1])

    with main_l:
        month_start = st.session_state.get("cal_calendar_month", date.today().replace(day=1))
        view_mode = st.session_state.get("cal_view_mode", "Month")
        if st.session_state.get("cal_selected_date_input"):
            st.session_state["cal_selected_date"] = st.session_state["cal_selected_date_input"]
        selected_date = st.session_state.get("cal_selected_date", date.today())
        if view_mode == "Week":
            grid_start = selected_date - timedelta(days=(selected_date.weekday() + 1) % 7); grid_end = grid_start + timedelta(days=6); row_count=1; days_per_row=7
        elif view_mode == "Day":
            grid_start = selected_date; grid_end = selected_date; row_count=1; days_per_row=1
        elif view_mode == "Agenda":
            grid_start = selected_date; grid_end = selected_date + timedelta(days=30); row_count=0; days_per_row=7
        else:
            grid_start = month_start - timedelta(days=(month_start.weekday() + 1) % 7); grid_end = grid_start + timedelta(days=41); row_count=6; days_per_row=7
        if selected_date.month != month_start.month and view_mode == "Month":
            selected_date = month_start
            st.session_state["cal_selected_date"] = selected_date
        elif selected_date.month != month_start.month and view_mode in {"Week", "Day", "Agenda"}:
            month_start = selected_date.replace(day=1)
            st.session_state["cal_calendar_month"] = month_start
        event_rows = df("SELECT e.*, p.name AS project_name, p.team AS project_team FROM calendar_events e LEFT JOIN projects p ON p.id=e.project_id WHERE date(e.end_date)>=date(?) AND date(e.start_date)<=date(?) ORDER BY e.start_date,e.start_time", (grid_start.isoformat(), grid_end.isoformat()))
        task_rows = df("SELECT t.id,t.name,t.start_date,t.due_date,t.status,t.team,p.team AS project_team,p.name AS project_name,t.assignee FROM tasks t LEFT JOIN projects p ON p.id=t.project_id WHERE date(t.due_date)>=date(?) AND date(t.due_date)<=date(?) ORDER BY t.due_date", (grid_start.isoformat(), grid_end.isoformat()))
        day_events = {}
        for _, ev in event_rows.iterrows():
            try:
                if cal_type_f != "All Event Types" and str(ev.event_type or "") != cal_type_f: continue
                if cal_team_f != "All Teams" and str(ev.get("project_team") or "") != cal_team_f: continue
                if cal_res_f != "All Resources" and cal_res_f not in (json.loads(ev.resources or "[]") if str(ev.resources or "").strip().startswith("[") else []): continue
                d0 = date.fromisoformat(str(ev.start_date)[:10]); d1 = date.fromisoformat(str(ev.end_date)[:10])
                cur = max(d0, grid_start)
                while cur <= min(d1, grid_end):
                    day_events.setdefault(cur, []).append((str(ev.title), f"{ev.project_name or 'No project'} · {ev.event_type or 'Event'}", "r" if "leave" in str(ev.event_type).lower() else ("b" if "meeting" in str(ev.event_type).lower() else "g")))
                    cur += timedelta(days=1)
            except Exception: pass
        for _, task in task_rows.iterrows():
            try:
                if cal_proj_f != "All Projects" and str(task.project_name or "") != cal_proj_f: continue
                if cal_team_f != "All Teams" and str(task.team or "") != cal_team_f: continue
                if cal_res_f != "All Resources" and str(task.assignee or "") != cal_res_f: continue
                if cal_type_f not in {"All Event Types", "Project Assignment"}: continue
                due = date.fromisoformat(str(task.due_date)[:10])
                day_events.setdefault(due, []).append((str(task.name), f"{task.project_name or 'Task'} · {task.status}", "g" if str(task.status) == "Completed" else ("r" if str(task.status) == "Overdue" else "b")))
            except Exception: pass
        ev_box_tones = {"g": ("#DCFCE7", "#065F46"), "b": ("#DBEAFE", "#1D4ED8"), "r": ("#FEE2E2", "#B91C1C")}
        cal_rows_html = ""
        cur_day = grid_start
        for _week in range(row_count):
            row_cells = ""
            for _ in range(days_per_row):
                in_m = cur_day.month == month_start.month
                is_sel = cur_day == selected_date
                bdr = "2px solid #2563EB" if is_sel else "1px solid #EDF2F7"
                bg = "#EFF6FF" if is_sel else ("#FFFFFF" if in_m else "#F8FAFC")
                d_col = "#0F172A" if in_m else "#94A3B8"
                pills = ""
                for etitle, esub, etone in day_events.get(cur_day, [])[:3]:
                    if cal_search_q and cal_search_q.lower() not in (etitle + " " + esub).lower(): continue
                    if cal_proj_f != "All Projects" and cal_proj_f.lower() not in esub.lower(): continue
                    pbg, pfg = ev_box_tones[etone]
                    pills += f'<div style="background:{pbg};color:{pfg};border-radius:6px;padding:4px 6px;font-size:10px;margin-top:4px;line-height:1.25;overflow:hidden"><div style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{html.escape(etitle)}</div><div style="font-size:9.5px;opacity:.9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{html.escape(esub)}</div></div>'
                row_cells += f'<div style="min-height:104px;border:{bdr};background:{bg};padding:6px 7px"><div style="font-size:12px;font-weight:700;color:{d_col}">{cur_day.day}</div>{pills}</div>'
                cur_day += timedelta(days=1)
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

        # Bottom summary cards are derived from the live records and current filters.
        bc1,bc2,bc3,bc4=st.columns(4)
        with bc1,st.container(border=True):
            section_title("Resources",f"{len(R)} loaded")
            if R.empty: st.caption("No resources imported yet.")
            else: st.dataframe(R[[c for c in ["name","team","status"] if c in R.columns]].head(6),hide_index=True,use_container_width=True,height=180)
        with bc2,st.container(border=True):
            section_title("Event Types")
            if event_rows.empty: st.caption("No calendar events in this month.")
            else: st.dataframe(event_rows["event_type"].fillna("Event").value_counts().rename_axis("Event Type").reset_index(name="Count"),hide_index=True,use_container_width=True,height=180)
        with bc3,st.container(border=True):
            section_title("Projects",f"{len(P)} loaded")
            if P.empty: st.caption("No projects imported yet.")
            else: st.dataframe(P[[c for c in ["name","status","progress"] if c in P.columns]].head(6),hide_index=True,use_container_width=True,height=180)
        with bc4,st.container(border=True):
            section_title("Calendar Legend")
            st.markdown(f"<div style='font-size:12px;line-height:2.1'><div><span style='color:{GREEN}'>●</span> Project / task</div><div><span style='color:{BLUE}'>●</span> Meeting / event</div><div><span style='color:{RED}'>●</span> Leave</div><div><span style='display:inline-block;width:10px;height:10px;border:2px solid #2563EB;border-radius:3px'></span> Selected date</div></div>",unsafe_allow_html=True)

    # --- RIGHT SIDE INSPECTOR PANEL (selected date, live events/tasks/resources) ---
    with main_r, st.container(border=True):
        selected_date = st.date_input("Selected Date", value=st.session_state.get("cal_selected_date", date.today()), key="cal_selected_date_input")
        st.session_state["cal_selected_date"] = selected_date
        if selected_date.month != st.session_state.get("cal_calendar_month", date.today().replace(day=1)).month:
            st.session_state["cal_calendar_month"] = selected_date.replace(day=1)
            st.rerun()
        st.markdown(f"<div style='font-size:18px;font-weight:800;color:#0F172A;margin:4px 0 8px 0'>{selected_date.strftime('%A, %b %d, %Y').replace(' 0',' ')}</div>", unsafe_allow_html=True)
        selected_tasks = T[(T["due_date"].astype(str).str[:10] == selected_date.isoformat())] if not T.empty else pd.DataFrame()
        selected_events = df("SELECT e.*, p.name AS project_name FROM calendar_events e LEFT JOIN projects p ON p.id=e.project_id WHERE date(start_date)<=date(?) AND date(end_date)>=date(?) ORDER BY start_time", (selected_date.isoformat(), selected_date.isoformat()))
        insp_tabs = st.tabs([f"Day ({len(selected_tasks)+len(selected_events)})", f"Team ({len(R)})", f"Projects ({len(P)})"])
        with insp_tabs[0]:
            if selected_events.empty and selected_tasks.empty:
                st.info("No scheduled events or tasks for this date.")
            for _, ev in selected_events.iterrows():
                st.markdown(f"**{html.escape(str(ev.title))}**  ")
                st.caption(f"{ev.start_time or 'All day'} – {ev.end_time or 'All day'} · {ev.event_type or 'Event'} · {ev.project_name or 'No project'}")
                if ev.description: st.write(str(ev.description))
                st.divider()
            for _, tr in selected_tasks.iterrows():
                st.markdown(f"**{html.escape(str(tr['name']))}**")
                st.caption(f"Task due · {tr.get('project','No project')} · {tr.get('status','Not Started')} · {tr.get('assignee') or 'Unassigned'}")
        with insp_tabs[1]:
            if R.empty:
                st.info("No resources loaded. Import a workbook in Settings.")
            else:
                res_view = R[[c for c in ["name", "team", "status", "capacity"] if c in R.columns]].copy()
                st.dataframe(res_view, hide_index=True, use_container_width=True, height=220)
        with insp_tabs[2]:
            if P.empty:
                st.info("No projects loaded. Import a workbook in Settings.")
            else:
                pr_view = P[[c for c in ["name", "status", "progress", "end_date", "days_left"] if c in P.columns]].copy()
                st.dataframe(pr_view, hide_index=True, use_container_width=True, height=220)
        st.markdown("<hr style='margin:8px 0 12px 0;border-color:#E2E8F0'><div style='font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:8px'>Quick Actions</div>", unsafe_allow_html=True)
        qa1, qa2, qa3, qa4 = st.columns(4)
        with qa1:
            if st.button("📅 Add Event", key="qa_add_ev", use_container_width=True):
                st.session_state.cal_modal = "add_event"; st.rerun()
        with qa2:
            if st.button("👥 View Team", key="qa_v_team", use_container_width=True):
                st.session_state.page = "Resources"; st.rerun()
        with qa3:
            if st.button("🕒 Set Leave", key="qa_set_lv", use_container_width=True):
                st.session_state.cal_modal = "add_event"; st.session_state["ae_type"] = "🔴 Leave"; st.rerun()
        with qa4:
            if st.button("⚙ Availability", key="qa_ed_av", use_container_width=True):
                st.session_state.page = "Resources"; st.rerun()

    if st.session_state.get("cal_modal") == "add_event":
        add_event_modal()
        
# ===========================================================================
# PAGE 7: REPORTS (Exact Replica of All 7 Report Mockup Screens)
# ===========================================================================
elif page == "Reports":
    if "report_tab" not in st.session_state:
        st.session_state.report_tab = "Overview"
    report_tabs = ["Overview", "Projects", "Resources", "Tasks", "Milestones", "Workload", "Custom"]
    active_rtab = st.session_state.report_tab if st.session_state.report_tab in report_tabs else "Overview"
    st.markdown(f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.1'>Reports</div><div style='font-size:13px;color:#64748B;margin-top:2px'>Live reports from the current Excel-fed project workspace.</div>", unsafe_allow_html=True)
    rhead, rproj, rteam, rexp = st.columns([2.2,1,1,0.8])
    with rhead:
        report_dates = st.date_input("Date Range", (date.today(), date.today()+timedelta(days=60)), key="rep_top_dates")
    with rproj:
        report_project = st.selectbox("Project", ["All Projects"]+projects_list, key="rep_top_proj")
    with rteam:
        report_team = st.selectbox("Team", ["All Teams"]+teams_list, key="rep_top_team")
    with rexp:
        st.download_button("⬇ Export", excel_bytes({"Projects":P,"Tasks":T,"Resources":R,"Milestones":M}), file_name=f"vektor_{active_rtab.lower()}_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True, key="rep_top_export_btn")
    nav_cols = st.columns(len(report_tabs))
    for col, name in zip(nav_cols, report_tabs):
        with col:
            if st.button(name, key=f"rep_nav_btn_{name}", type="primary" if active_rtab == name else "secondary", use_container_width=True):
                st.session_state.report_tab = name
                st.rerun()
    projects_view=P.copy(); tasks_view=T.copy(); resources_view=R.copy(); milestones_view=M.copy()
    if report_project != "All Projects":
        projects_view=projects_view[projects_view["name"]==report_project]
        tasks_view=tasks_view[tasks_view["project"]==report_project] if "project" in tasks_view else tasks_view.iloc[0:0]
        milestones_view=milestones_view[milestones_view["project"]==report_project] if "project" in milestones_view else milestones_view.iloc[0:0]
    if report_team != "All Teams":
        if "team" in projects_view: projects_view=projects_view[projects_view["team"]==report_team]
        if "team" in tasks_view: tasks_view=tasks_view[tasks_view["team"]==report_team]
        if "team" in resources_view: resources_view=resources_view[resources_view["team"]==report_team]
        if "team" in milestones_view: milestones_view=milestones_view[milestones_view["team"]==report_team]
    if isinstance(report_dates, tuple) and len(report_dates)==2:
        lo,hi=report_dates
        if "due_date" in tasks_view:
            due=pd.to_datetime(tasks_view["due_date"],errors="coerce")
            tasks_view=tasks_view[due.isna() | ((due.dt.date>=lo)&(due.dt.date<=hi))]
        if "due_date" in milestones_view:
            due=pd.to_datetime(milestones_view["due_date"],errors="coerce")
            milestones_view=milestones_view[due.isna() | ((due.dt.date>=lo)&(due.dt.date<=hi))]
    if active_rtab == "Custom":
        with st.container(border=True):
            st.markdown("**Build Custom Report**")
            cr_type=st.selectbox("Report Type",["Projects","Tasks","Resources","Milestones"],key="cr_rep_type")
            cr_group=st.selectbox("Group By",["Status","Team","Priority","Project","Assignee"],key="cr_group_by")
            cr_metrics=st.multiselect("Metrics",["Count","Progress","Days Left"],default=["Count","Progress"],key="cr_metrics")
            source={"Projects":projects_view,"Tasks":tasks_view,"Resources":resources_view,"Milestones":milestones_view}[cr_type]
            group_col={"Status":"status","Team":"team","Priority":"priority","Project":"project","Assignee":"assignee"}[cr_group]
            if source.empty: st.info("No data matches the selected filters.")
            elif group_col not in source.columns: st.info(f"{cr_group} is not available for this data type.")
            else:
                agg={}
                source_for_group=source.copy()
                source_for_group["_Record Count"]=1
                if "Count" in cr_metrics: agg["_Record Count"]="sum"
                if "Progress" in cr_metrics and "progress" in source_for_group: agg["progress"]="mean"
                if "Days Left" in cr_metrics and "days_left" in source_for_group: agg["days_left"]="mean"
                if not agg: agg={"_Record Count":"sum"}
                grouped=source_for_group.groupby(group_col,dropna=False).agg(agg).reset_index()
                st.dataframe(grouped,hide_index=True,use_container_width=True)
                st.download_button("Download Custom Report",excel_bytes({"Custom Report":grouped}),file_name="vektor_custom_report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="cr_download")
    else:
        source_map={"Projects":projects_view,"Resources":resources_view,"Tasks":tasks_view,"Milestones":milestones_view,"Workload":resources_view,"Overview":projects_view}
        source=source_map[active_rtab]
        a,b,c,d=st.columns(4)
        with a: metric_card("Projects",len(projects_view),"Current filter","📁","blue")
        with b: metric_card("Tasks",len(tasks_view),f"{int((tasks_view.status=='Completed').sum()) if not tasks_view.empty else 0} completed","☑","green")
        with c: metric_card("At Risk / Overdue",int(tasks_view.status.isin(["At Risk","Overdue"]).sum()) if not tasks_view.empty else 0,"Date-driven","⚠","orange")
        with d: metric_card("Resources",len(resources_view),"Loaded owners","👥","purple")
        left,right=st.columns([1.1,1])
        with left, st.container(border=True):
            section_title(f"{active_rtab} by Status")
            if not source.empty and "status" in source.columns:
                chart_status_donut(source,"status",center_label=active_rtab,height=250)
            else: st.info("No status data available for this report.")
        with right, st.container(border=True):
            section_title("Project Completion")
            if not projects_view.empty:
                progress_bar_chart(projects_view.rename(columns={"name":"Project"}),name="Project",val="progress",height=250)
            else: st.info("No projects match the selected filters.")
        with st.container(border=True):
            section_title(f"{active_rtab} Records", f"{len(source)} records")
            if source.empty: st.info("No records match the current filters.")
            else:
                st.dataframe(source,hide_index=True,use_container_width=True,height=360)
                st.download_button("Download Filtered Records",excel_bytes({active_rtab:source}),file_name=f"vektor_{active_rtab.lower()}_filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="rep_filtered_download")
    if st.button("Save Current Report", key="save_current_report"):
        cfg={"tab":active_rtab,"project":report_project,"team":report_team,"dates":[str(x) for x in report_dates] if isinstance(report_dates,tuple) else []}
        sql("INSERT INTO saved_reports(name,report_type,config,last_generated,created_by) VALUES(?,?,?,?,?)",(f"{active_rtab} Report",active_rtab,json.dumps(cfg),date.today().isoformat(),"Current User"))
        st.success("Report configuration saved.")


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

    # Excel is the operational update source; uploads replace live projects/tasks/resources.
    with st.container(border=True):
        st.markdown("**Excel Data Source**")
        st.caption("Upload the portfolio workbook to refresh the app. The Portfolio sheet is imported; the workbook's other sheets are retained in the source file and are not used as transaction data.")
        imp_col, exp_col = st.columns([2, 1])
        with imp_col:
            portfolio_upload = st.file_uploader("Upload Excel portfolio", type=["xlsx", "xlsm"], key="settings_portfolio_upload")
            if portfolio_upload is not None and st.button("Import / Refresh App Data", type="primary", key="settings_import_portfolio"):
                try:
                    counts = import_portfolio_workbook(portfolio_upload)
                    st.success(f"Imported {counts['projects']} programs, {counts['tasks']} items, and {counts['resources']} owners.")
                    st.session_state["_vektor_has_real_data"] = True
                    st.rerun()
                except Exception as exc:
                    st.error(f"Import failed: {exc}")
        with exp_col:
            st.download_button("Download Current App Data", data=current_data_workbook(), file_name="Vektor_Current_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="settings_export_data")
            st.caption("Edit the downloaded workbook and upload it again to refresh the app.")

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
            logo_file = st.file_uploader("Workspace Logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key="settings_workspace_logo")
            if logo_file is not None:
                put_setting("workspace_logo_b64", base64.b64encode(logo_file.getvalue()).decode("ascii"))
                st.image(logo_file, width=180)

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
            st.selectbox("Date Format", ["MMM DD, YYYY", "DD/MM/YYYY", "MM/DD/YYYY"], index=0, key="set_ws_dfmt")
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

        um_df = R.copy()
        if um_q:
            qlow = um_q.lower()
            um_df = um_df[um_df.apply(lambda row: any(qlow in str(v).lower() for v in row.values), axis=1)]
        if um_df.empty:
            st.info("No users/resources match the search.")
        else:
            display_users = um_df[[c for c in ["id", "name", "email", "role", "team", "status"] if c in um_df.columns]].copy()
            display_users = display_users.rename(columns={"id":"ID", "name":"Name", "email":"Email", "role":"Role", "team":"Team", "status":"Status"})
            st.dataframe(display_users, hide_index=True, use_container_width=True, height=280)
            user_action_cols = st.columns([1.5, 1.5, 1.5, 2.5])
            with user_action_cols[0]:
                user_choice = st.selectbox("Select user", um_df["name"].tolist(), key="settings_user_action_target", label_visibility="collapsed")
            with user_action_cols[1]:
                if st.button("Set Available", key="settings_user_available", use_container_width=True):
                    sql("UPDATE resources SET status='Available' WHERE name=?", (user_choice,)); st.rerun()
            with user_action_cols[2]:
                if st.button("Set On Leave", key="settings_user_leave", use_container_width=True):
                    sql("UPDATE resources SET status='On Leave', capacity=0 WHERE name=?", (user_choice,)); st.rerun()
            with user_action_cols[3]:
                if st.button("Delete Selected User", key="settings_user_delete", use_container_width=True):
                    sql("DELETE FROM resources WHERE name=?", (user_choice,)); st.rerun()

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
