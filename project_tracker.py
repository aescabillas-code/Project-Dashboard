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
import hashlib
from typing import Iterable
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
    names = ("projects", "tasks", "resources", "milestones", "settings", "saved_reports", "audit_log")
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
    tables = ("projects", "tasks", "resources", "milestones", "settings", "saved_reports", "audit_log")
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
    """Remove seeded mock records when the first real project/task/resource is created."""
    if st.session_state.get("_vektor_has_real_data"):
        return
    tables = ("projects", "tasks", "resources", "milestones", "saved_reports")
    with conn() as c:
        for table in tables:
            if table != except_table:
                c.execute(f"DELETE FROM {table}")
        c.commit()
    st.session_state["_vektor_has_real_data"] = True


def _recalculate_project_rollups() -> None:
    """Keep project progress, task counts, health, and status synchronized with task changes."""
    with conn() as c:
        projects = c.execute("SELECT id, start_date, end_date, status FROM projects").fetchall()
        today = date.today()
        for project in projects:
            tasks = c.execute(
                "SELECT status, progress, due_date FROM tasks WHERE project_id=?", (project["id"],)
            ).fetchall()
            total = len(tasks)
            done = sum(1 for task in tasks if str(task["status"] or "").lower() == "completed")
            progress = round(sum(max(0, min(100, int(task["progress"] or 0))) for task in tasks) / total) if total else 0
            end_date = None
            try:
                end_date = date.fromisoformat(str(project["end_date"])[:10])
            except Exception:
                pass
            if total and done == total:
                status, health = "Completed", "Healthy"
            elif end_date and end_date < today and progress < 100:
                status, health = "Delayed", "Critical"
            elif end_date and (end_date - today).days <= 7 and progress < 75:
                status, health = "At Risk", "At Risk"
            elif total and progress > 0:
                status, health = "In Progress", "Healthy"
            else:
                status, health = (project["status"] or "Not Started"), "Healthy"
            c.execute(
                "UPDATE projects SET tasks_done=?, tasks_total=?, progress=?, days_left=?, status=?, health=? WHERE id=?",
                (done, total, progress, (end_date - today).days if end_date else None, status, health, project["id"]),
            )
        c.commit()


def sql(q: str, p: tuple = ()) -> None:
    """Execute an app mutation, then persist it and recalculate dependent project metrics."""
    normalized = " ".join(q.strip().lower().split())
    match = re.search(r"\b(?:insert\s+into|update|delete\s+from)\s+([a-z_]+)", normalized)
    table = match.group(1) if match else ""
    is_entity_mutation = table in {"projects", "tasks", "resources", "milestones", "saved_reports", "settings", "audit_log"}
    is_insert = normalized.startswith("insert into")
    # First user-created record marks the end of demo-only mode. Keep existing projects
    # available when the first real task/resource is added so foreign references still work.
    if is_entity_mutation and is_insert and not st.session_state.get("_vektor_has_real_data"):
        _clear_mock_records(except_table="projects" if table != "projects" else "")
    with conn() as c:
        c.execute(q, p)
        c.commit()
    _cached_df.clear()
    if table in {"tasks", "projects", "milestones"}:
        _recalculate_project_rollups()
    if is_entity_mutation:
        _persist_snapshot()


@st.cache_data(show_spinner=False, ttl=30)
def _cached_df(q: str, p: tuple, db_mtime: float) -> pd.DataFrame:
    with conn() as c:
        return pd.read_sql_query(q, c, params=p)

def df(q: str, p: tuple = ()) -> pd.DataFrame:
    try:
        db_mtime = os.path.getmtime(DB)
    except OSError:
        db_mtime = 0.0
    return _cached_df(q, p, db_mtime).copy()


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
                id INTEGER PRIMARY KEY,
                event TEXT,
                item_type TEXT,
                item_id INTEGER,
                details TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )

        if False and c.execute("SELECT COUNT(*) FROM projects").fetchone()[0] == 0:
            # Demo seeding disabled: production data comes from Excel or user entry.
            ps = [
                (
                    1, "Website Revamp", "Alice Morgan", "Marketing", "Web Development",
                    "On Track", "High", "2024-04-24", "2024-06-30", 70, 45, 14, 20,
                    "Web,Marketing,Corporate", "#00875A", "Redesign corporate website",
                    "Redesign and launch the corporate website with improved user experience.", "Healthy"
                ),
                (
                    2, "E-Commerce Platform", "Brian Taylor", "IT", "IT & Development",
                    "In Progress", "High", "2024-03-15", "2024-06-15", 50, 30, 12, 24,
                    "IT,Development", "#1473E6", "Build new e-commerce platform",
                    "Build new e-commerce platform to expand digital presence.", "Healthy"
                ),
                (
                    3, "Mobile App Development", "Emily Rodriguez", "Design", "Mobile Development",
                    "At Risk", "Medium", "2024-04-01", "2024-08-30", 30, 90, 8, 26,
                    "Mobile,Customer", "#7C3AED", "Develop customer mobile app",
                    "Develop customer mobile application for iOS and Android.", "At Risk"
                ),
                (
                    4, "Data Migration", "Carlos Diaz", "IT", "Cloud & Data",
                    "On Track", "High", "2024-05-01", "2024-07-31", 80, 60, 18, 22,
                    "Data,Cloud", "#F28C00", "Migrate legacy data to cloud",
                    "Migrate legacy data to cloud infrastructure with minimal downtime.", "Healthy"
                ),
                (
                    5, "IT Infrastructure Upgrade", "Dana Scott", "IT", "Infrastructure",
                    "At Risk", "High", "2024-05-20", "2024-09-15", 20, 110, 6, 28,
                    "Infrastructure,IT", "#00875A", "Upgrade servers and network",
                    "Upgrade servers and network infrastructure to support business growth.", "At Risk"
                ),
                (
                    6, "Security Enhancement", "Mark Wilson", "IT", "Security & Compliance",
                    "On Track", "Medium", "2024-04-10", "2024-07-10", 65, 50, 13, 20,
                    "Security,Compliance", "#00875A", "Implement security measures",
                    "Implement security measures to strengthen system protection and compliance.", "Healthy"
                ),
                (
                    7, "Customer Portal", "Sarah Kim", "Marketing", "Web Development",
                    "On Track", "High", "2024-03-01", "2024-06-30", 90, 45, 18, 20,
                    "Customer,Web", "#7C3AED", "Enhance customer portal",
                    "Enhance customer portal with new features and improved user experience.", "Healthy"
                ),
                (
                    8, "AI Integration", "Kevin Lee", "IT", "AI & Innovation",
                    "Not Started", "Medium", "2024-05-05", "2024-10-31", 15, 140, 3, 18,
                    "AI,Innovation", "#EC4899", "Integrate AI capabilities",
                    "Integrate AI capabilities to improve business processes and customer support.", "Healthy"
                ),
                (
                    9, "Process Automation", "Jason Tan", "Operations", "Operations",
                    "In Progress", "Low", "2024-06-01", "2024-09-30", 40, 120, 8, 20,
                    "Automation,Operations", "#1473E6", "Automate internal processes",
                    "Automate internal processes to increase efficiency and reduce manual work.", "Healthy"
                ),
                (
                    10, "Training Program", "Nicole Reyes", "Operations", "HR & Training",
                    "In Progress", "Low", "2024-04-15", "2024-08-15", 55, 75, 11, 20,
                    "HR,Training", "#DB2777", "Employee training and upskilling",
                    "Employee training and upskilling program for new systems.", "Healthy"
                ),
                (
                    11, "Vendor Management System", "David Park", "Operations", "Procurement",
                    "Not Started", "Medium", "2024-05-10", "2024-11-30", 10, 170, 2, 18,
                    "Procurement,IT", "#0284C7", "Implement vendor management",
                    "Implement vendor management system to streamline procurement.", "At Risk"
                ),
                (
                    12, "Sustainability Initiative", "Laura Torres", "Operations", "Sustainability",
                    "Not Started", "Low", "2024-06-15", "2024-12-31", 5, 200, 1, 24,
                    "Sustainability,ESG", "#00875A", "Implement green IT practices",
                    "Implement green IT practices and sustainability programs.", "Critical"
                ),
            ]
            c.executemany(
                """
                INSERT INTO projects(
                    id,name,owner,team,category,status,priority,start_date,end_date,
                    progress,days_left,tasks_done,tasks_total,tags,bar_color,short_desc,description,health
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                ps,
            )

            people = [
                "Alice Morgan", "Brian Taylor", "Emily Rodriguez", "Carlos Diaz", "Dana Scott",
                "Mark Wilson", "Sarah Kim", "Kevin Lee", "Jason Tan", "Nicole Reyes",
                "David Park", "Laura Torres", "James Miller", "Ryan Cruz", "Michael Tan",
                "Daniel Perez", "Ella Ramirez", "Chris Padilla", "Anna Reyes", "Kevin Santos",
                "Lisa Chan", "David Lim", "Maria Santos", "John Cruz",
            ]
            teams_list_seed = [
                "Marketing", "IT", "Design", "IT", "IT",
                "IT", "Marketing", "IT", "Operations", "Operations",
                "Operations", "Operations", "IT", "Operations", "IT",
                "Operations", "Design", "IT", "Marketing", "Operations",
                "Design", "IT", "Marketing", "Others",
            ]
            for i, (n, t) in enumerate(zip(people, teams_list_seed)):
                status = "On Leave" if i in (14, 20) else ("Partial" if i in (2, 6, 11, 17) else "Available")
                role = "Admin" if i in (0, 12) else ("Project Manager" if i < 12 else "Team Member")
                cap = 0 if status == "On Leave" else (50 if status == "Partial" else 100)
                c.execute(
                    "INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",
                    (n, n.lower().replace(" ", ".") + "@vektor.example", t, role, status, cap, ""),
                )

            tasks_seed = [
                ("UI/UX Design Complete", 1, "Design", "Completed", "High", "2024-04-01", "2024-04-30", "Alice Morgan", 100, "", "Design phase completed."),
                ("Homepage Development", 1, "IT", "In Progress", "High", "2024-04-10", "2024-05-15", "James Miller", 65, "UI/UX Design Complete", "Frontend responsive build."),
                ("Content Review", 7, "Marketing", "In Progress", "Medium", "2024-04-15", "2024-05-20", "Sarah Kim", 55, "", "Copywriting and creative review."),
                ("Infrastructure Setup", 5, "IT", "Not Started", "High", "2024-05-01", "2024-05-15", "Dana Scott", 0, "Requirements Gathering", "Provisioning cluster nodes."),
                ("Sustainability Plan", 12, "Operations", "Completed", "Medium", "2024-05-05", "2024-05-30", "Laura Torres", 100, "", "Charter approved by leadership."),
                ("Process Documentation", 9, "Operations", "In Progress", "Low", "2024-05-10", "2024-06-20", "Jason Tan", 50, "", "Mapping core operational workflows."),
                ("Training Materials", 10, "Operations", "Not Started", "Medium", "2024-05-20", "2024-06-30", "Nicole Reyes", 0, "Content Review", "Sales enablement deck."),
                ("QA Testing", 1, "IT", "In Progress", "High", "2024-05-25", "2024-06-25", "Chris Padilla", 35, "Homepage Development", "Cross-browser regression testing."),
                ("API Integration", 2, "IT", "In Progress", "High", "2024-04-12", "2024-05-25", "Brian Taylor", 60, "", "Payment gateway & cart API."),
                ("UI Development", 3, "Design", "In Progress", "Medium", "2024-04-18", "2024-06-10", "Emily Rodriguez", 40, "", "Mobile screens implementation."),
                ("Data Mapping", 4, "IT", "Completed", "Low", "2024-04-01", "2024-04-25", "Carlos Diaz", 100, "", "Schema validation finished."),
            ]
            c.executemany(
                "INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                tasks_seed,
            )

            ms_seed = [
                ("Project Kick-off", 1, "Marketing", "Completed", "2024-04-20", "2024-04-24", "2024-04-24", "", "Initial kick-off complete."),
                ("Requirements Finalized", 1, "Marketing", "Completed", "2024-04-25", "2024-05-05", "2024-05-05", "Project Kick-off", "Scope signed off."),
                ("Design Sign-off", 1, "Design", "Completed", "2024-04-10", "2024-04-30", "2024-04-30", "Requirements Finalized", "Design phase completed."),
                ("Development Complete", 1, "IT", "In Progress", "2024-05-01", "2024-05-31", "", "Design Sign-off", "Core build underway."),
                ("UAT Completion", 1, "IT", "Not Started", "2024-06-01", "2024-06-15", "", "Development Complete", "Pending staging build."),
                ("Go-Live", 1, "IT", "Not Started", "2024-06-16", "2024-06-30", "", "UAT Completion", "Production launch."),
                ("MVP Release", 3, "Design", "At Risk", "2024-04-15", "2024-05-15", "", "UI Development", "Awaiting final assets."),
                ("Data Migration Cutover", 4, "IT", "On Track", "2024-05-10", "2024-06-15", "", "Data Mapping", "Environment setup ongoing."),
                ("Infrastructure Go-Live", 5, "IT", "On Track", "2024-06-01", "2024-07-01", "", "Infrastructure Setup", "Scheduled cutover."),
            ]
            c.executemany(
                "INSERT INTO milestones(name,project_id,team,status,start_date,due_date,completion_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?)",
                ms_seed,
            )

            default_settings = [
                ("workspace_name", "Vektor"),
                ("workspace_description", "Project management workspace for planning, tracking, and delivery."),
                ("theme", "HPE Green"),
                ("language", "English"),
                ("date_format", "MMM DD, YYYY"),
                ("default_landing", "Overview"),
                ("default_project_status", "On Track"),
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
                ("integration_Microsoft 365", "true"),
                ("integration_Slack", "true"),
                ("integration_Microsoft Teams", "false"),
                ("integration_OneDrive", "false"),
                ("audit_logs", "true"),
                ("data_export", "true"),
                ("automatic_backups", "true"),
                ("maintenance_mode", "false"),
                ("session_timeout", "8 hours"),
            ]
            c.executemany("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", default_settings)

            if c.execute("SELECT COUNT(*) FROM saved_reports").fetchone()[0] == 0:
                saved_seed = [
                    ("Monthly Project Summary", "Project Report", json.dumps({"type": "Project Report", "group": "Status", "metrics": ["Project Status", "Progress"], "viz": "Bar Chart"}), "2024-04-20", "Alice Morgan"),
                    ("Q2 Resource Allocation", "Resource Report", json.dumps({"type": "Resource Report", "group": "Team", "metrics": ["Team Members"], "viz": "Column Chart"}), "2024-04-18", "Sarah Kim"),
                    ("Overdue Tasks & Milestones", "Task Report", json.dumps({"type": "Task Report", "group": "Priority", "metrics": ["Task Status", "Priority"], "viz": "Pie Chart"}), "2024-04-15", "Dana Scott"),
                ]
                c.executemany(
                    "INSERT INTO saved_reports(name,report_type,config,last_generated,created_by) VALUES(?,?,?,?,?)",
                    saved_seed,
                )
            c.commit()



DEFAULT_SETTINGS = {
    "workspace_name": "Vektor", "workspace_description": "Project management workspace for planning, tracking, and delivery.",
    "theme": "HPE Green", "language": "English", "date_format": "MMM DD, YYYY", "default_landing": "Overview",
    "audit_logs": "true", "data_export": "true", "automatic_backups": "true", "maintenance_mode": "false",
    "session_timeout": "8 hours",
}

def ensure_default_settings() -> None:
    with conn() as c:
        c.executemany("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", list(DEFAULT_SETTINGS.items()))
        c.commit()

def _text(v: object) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()

def _iso(v: object) -> str:
    if v is None or _text(v) == "":
        return ""
    try:
        return pd.to_datetime(v, errors="raise").date().isoformat()
    except Exception:
        return ""

def _split_people(value: object) -> list[str]:
    raw = _text(value)
    if not raw:
        return []
    names = re.split(r"\s*(?:&|/|;|,|\band\b)\s*", raw, flags=re.I)
    return [n.strip() for n in names if n.strip()]

def extract_portfolio_workbook(uploaded_file) -> dict[str, list[dict]]:
    """Read the Portfolio worksheet and normalize it for the app database."""
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    book = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    portfolio_name = next((n for n in book.sheet_names if n.strip().lower() == "portfolio"), None)
    if not portfolio_name:
        raise ValueError("The workbook must contain a worksheet named 'Portfolio'.")
    preview = pd.read_excel(io.BytesIO(raw), sheet_name=portfolio_name, header=None, engine="openpyxl", nrows=25)
    header_row = None
    for idx, row in preview.iterrows():
        vals = {_text(v).lower() for v in row.tolist()}
        if "id" in vals and "program" in vals and any("initiative" in v for v in vals):
            header_row = idx
            break
    if header_row is None:
        raise ValueError("Could not find the Portfolio header row. Required columns include ID, Program, Initiative / Milestone / Task, Owner, and Status.")
    data = pd.read_excel(io.BytesIO(raw), sheet_name=portfolio_name, header=header_row, engine="openpyxl")
    data.columns = [_text(c) or f"Column_{i}" for i, c in enumerate(data.columns)]
    def col(*names):
        normalized = {re.sub(r"\W+", " ", c.lower()).strip(): c for c in data.columns}
        for name in names:
            key = re.sub(r"\W+", " ", name.lower()).strip()
            if key in normalized: return normalized[key]
        return None
    c_id, c_program = col("ID"), col("Program")
    c_item = col("Initiative / Milestone / Task", "Initiative Milestone Task")
    c_start, c_end, c_owner, c_status = col("Start Date"), col("End Date"), col("Owner"), col("Status")
    c_update, c_remarks, c_pc = col("Update"), col("Remarks"), col("PC")
    required = {"ID": c_id, "Program": c_program, "Initiative / Milestone / Task": c_item}
    missing = [k for k,v in required.items() if not v]
    if missing: raise ValueError("Missing required Portfolio columns: " + ", ".join(missing))
    rows=[]
    for _,r in data.iterrows():
        rid=_text(r.get(c_id)); program=_text(r.get(c_program)); item=_text(r.get(c_item))
        if not rid or not item: continue
        rows.append({"source_id":rid,"program":program,"item":item,"start":_iso(r.get(c_start)) if c_start else "","end":_iso(r.get(c_end)) if c_end else "","owner":_text(r.get(c_owner)) if c_owner else "","status":_text(r.get(c_status)) or "Not Started","update":_text(r.get(c_update)) if c_update else "","remarks":_text(r.get(c_remarks)) if c_remarks else "","pc":_text(r.get(c_pc)) if c_pc else ""})
    if not rows: raise ValueError("No extractable rows were found in the Portfolio worksheet.")
    groups={}
    for r in rows:
        root=r["source_id"].split(".")[0]
        groups.setdefault(root,[]).append(r)
    projects=[]; tasks=[]; milestones=[]; resources={}
    for root,items in groups.items():
        base=next((x for x in items if x["source_id"]==root),items[0])
        leaf=[x for x in items if x["source_id"]!=root] or items
        statuses=[x["status"].lower() for x in leaf]
        complete=sum(1 for x in statuses if x in {"completed","complete","done"})
        progress=round(100*complete/len(leaf)) if leaf else 0
        starts=[x["start"] for x in items if x["start"]]; ends=[x["end"] for x in items if x["end"]]
        owner=base["owner"] or next((x["owner"] for x in items if x["owner"]),"")
        status=base["status"] or ("Completed" if progress==100 else "In Progress" if progress else "Not Started")
        project={"source_id":root,"name":base["program"] or base["item"],"owner":owner,"team":base["pc"] or "Portfolio","category":base["program"] or "Portfolio","status":status,"priority":"Medium","start_date":min(starts) if starts else "","end_date":max(ends) if ends else "","progress":progress,"tasks_done":complete,"tasks_total":len(leaf),"tags":base["program"],"short_desc":base["item"],"description":base["remarks"] or base["update"] or base["item"],"health":"Critical" if status.lower() in {"delayed","overdue"} else "At Risk" if "risk" in status.lower() else "Healthy"}
        projects.append(project)
        for x in items:
            if x["source_id"]==root: continue
            rec={"source_id":x["source_id"],"project_source_id":root,"name":x["item"],"team":x["pc"] or base["pc"] or "Portfolio","status":x["status"],"priority":"Medium","start_date":x["start"],"due_date":x["end"],"assignee":x["owner"],"progress":100 if x["status"].lower() in {"completed","complete","done"} else 50 if "progress" in x["status"].lower() else 0,"dependency":"","notes":"\n".join(v for v in [x["update"],x["remarks"]] if v)}
            if "milestone" in x["item"].lower(): milestones.append(rec)
            else: tasks.append(rec)
        for x in items:
            for person in _split_people(x["owner"]): resources.setdefault(person,{"name":person,"email":"","team":x["pc"] or "Portfolio","role":"Team Member","status":"Available","capacity":100,"notes":"Imported from Portfolio workbook"})
    return {"projects":projects,"tasks":tasks,"milestones":milestones,"resources":list(resources.values()),"source_hash":hashlib.sha256(raw).hexdigest(),"rows":len(rows)}

def import_portfolio(payload: dict, replace_existing: bool = True) -> dict[str,int]:
    with conn() as c:
        if replace_existing:
            for table in ("tasks","milestones","projects","resources"):
                c.execute(f"DELETE FROM {table}")
        project_ids={}
        for p in payload["projects"]:
            cur=c.execute("INSERT INTO projects(name,owner,team,category,status,priority,start_date,end_date,progress,days_left,tasks_done,tasks_total,tags,bar_color,short_desc,description,health) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(p["name"],p["owner"],p["team"],p["category"],p["status"],p["priority"],p["start_date"],p["end_date"],p["progress"],None,p["tasks_done"],p["tasks_total"],p["tags"],GREEN,p["short_desc"],p["description"],p["health"]))
            project_ids[p["source_id"]]=cur.lastrowid
        for t in payload["tasks"]:
            c.execute("INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(t["name"],project_ids[t["project_source_id"]],t["team"],t["status"],t["priority"],t["start_date"],t["due_date"],t["assignee"],t["progress"],t["dependency"],t["notes"]))
        for m in payload["milestones"]:
            c.execute("INSERT INTO milestones(name,project_id,team,status,start_date,due_date,completion_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?)",(m["name"],project_ids[m["project_source_id"]],m["team"],m["status"],m["start_date"],m["due_date"],m["due_date"] if m["status"].lower()=="completed" else "",m["dependency"],m["notes"]))
        for r in payload["resources"]:
            c.execute("INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",(r["name"],r["email"],r["team"],r["role"],r["status"],r["capacity"],r["notes"]))
        c.execute("INSERT INTO settings(key,value) VALUES('last_import_hash',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",(payload["source_hash"],))
        c.execute("INSERT INTO audit_log(event,item_type,details) VALUES(?,?,?)",("Imported","Portfolio Workbook",f"{payload['rows']} rows"))
        c.commit()
    _cached_df.clear(); _recalculate_project_rollups(); _persist_snapshot()
    return {k:len(payload[k]) for k in ("projects","tasks","milestones","resources")}

def clear_selected_data(selections: Iterable[str]) -> None:
    mapping={"Projects":"projects","Tasks":"tasks","Resources":"resources","Milestones":"milestones","Saved reports":"saved_reports","Audit log":"audit_log","Settings":"settings"}
    selected=[mapping[x] for x in selections if x in mapping]
    with conn() as c:
        for table in selected: c.execute(f"DELETE FROM {table}")
        c.commit()
    if "settings" in selected: ensure_default_settings()
    _cached_df.clear(); _persist_snapshot()

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


init_db()
ensure_default_settings()
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
                    The {html.escape(r['name'])} project aims to deliver a modern, user-friendly corporate website with improved performance, enhanced user experience, and updated content. The project is currently on track with {pct}% completion. Key development activities are in progress and the team is aligned for the upcoming milestones.
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
            act_rows = [
                ("Apr 24, 2024 10:15 AM", "#00875A", "Project created", "Alice Morgan"),
                ("Apr 23, 2024 04:20 PM", "#1473E6", "Milestone status changed to In Progress", "Emily Rodriguez"),
                ("Apr 22, 2024 03:45 PM", "#1473E6", "Project details updated", "Alice Morgan"),
                ("Apr 20, 2024 11:10 AM", "#64748B", "Team member added", "Sarah Kim"),
            ]
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
            p_name = st.text_input("Project Name *", value=str(r["name"]) if r is not None else "Website Revamp")
        with f_b:
            default_desc = (
                "Redesign and launch the corporate website with improved user experience and modern design."
                if (r is not None and r["name"] == "Website Revamp")
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
            s_dt = date.fromisoformat(r.start_date) if r is not None and r.start_date else date(2024, 4, 24)
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
        st.date_input("Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="ov_date_sel", label_visibility="collapsed")

    st.write("")
    # 6 KPI Cards Row (Exact numbers & progress bars from Project Overview screenshot)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        kpi_card_mockup("Total Projects", "5", sub_text="Active Projects", icon_html="❖", icon_bg="#DCFCE7", icon_fg="#00875A")
    with k2:
        kpi_card_mockup("Total Tasks", "142", sub_text="Across All Projects", icon_html="✓", icon_bg="#DCFCE7", icon_fg="#00875A")
    with k3:
        kpi_card_mockup("Completed Tasks", "78", "(55%)", bar_pct=55, icon_html="◕", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    with k4:
        kpi_card_mockup("In Progress", "46", "(32%)", bar_pct=32, icon_html="▶", icon_bg="#DBEAFE", icon_fg="#1473E6", bar_color="#1473E6")
    with k5:
        kpi_card_mockup("On Hold", "12", "(8%)", bar_pct=8, icon_html="❚❚", icon_bg="#FEF3C7", icon_fg="#F28C00", bar_color="#F28C00")
    with k6:
        kpi_card_mockup("Overdue", "6", "(4%)", bar_pct=4, icon_html="✕", icon_bg="#FEE2E2", icon_fg="#EF4444", bar_color="#EF4444")

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

        ov_projs = P[P.id.isin([1, 2, 3, 4, 5])].copy()
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
            ts_df = pd.DataFrame(
                {
                    "Status": ["Completed", "In Progress", "On Hold", "Overdue"],
                    "Count": [78, 46, 12, 6],
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
                        text="<b style='font-size:20px;color:#0F172A'>142</b><br><span style='font-size:11px;color:#64748B'>Total Tasks</span>",
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
        ms_ov = [
            ("#00875A", "Design Sign-off", "Website Revamp", "Apr 30, 2024", 5, "On Track"),
            ("#F28C00", "MVP Release", "Mobile App Development", "May 15, 2024", 20, "At Risk"),
            ("#1473E6", "UAT Completion", "E-Commerce Platform", "May 30, 2024", 35, "On Track"),
            ("#7C3AED", "Data Migration Cutover", "Data Migration", "Jun 15, 2024", 50, "On Track"),
            ("#00875A", "Infrastructure Go-Live", "IT Infrastructure Upgrade", "Jul 1, 2024", 66, "On Track"),
        ]
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
                    "Count": [26, 34, 54, 28],
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
                        text="<b style='font-size:20px;color:#0F172A'>142</b><br><span style='font-size:11px;color:#64748B'>Total Tasks</span>",
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
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#EF4444">●</span> &nbsp;High</span><b>26</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;Medium</span><b>34</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#1473E6">●</span> &nbsp;Low</span><b>54</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;None</span><b>28</b></div>
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

        tw_rows = [
            ("Alice Morgan", 28, 14, 10, 4, 75),
            ("Brian Taylor", 24, 10, 12, 2, 60),
            ("Emily Rodriguez", 22, 12, 8, 2, 65),
            ("Carlos Diaz", 18, 8, 8, 2, 50),
            ("Dana Scott", 16, 8, 6, 2, 45),
        ]
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
        ra_rows = [
            ("Apr 24, 2024 10:15 AM", "#00875A", "Task marked as completed", "Website Revamp", "Homepage Design"),
            ("Apr 24, 2024 09:42 AM", "#1473E6", "Task updated", "E-Commerce Platform", "API Integration"),
            ("Apr 24, 2024 08:30 AM", "#1473E6", "New task created", "Mobile App Development", "UI Development"),
            ("Apr 23, 2024 04:20 PM", "#1473E6", "Status changed to In Progress", "Data Migration", "Data Mapping"),
            ("Apr 23, 2024 01:10 PM", "#1473E6", "Task assigned", "IT Infrastructure Upgrade", "Server Setup"),
        ]
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
        st.date_input("Range", (date(2024, 4, 24), date(2024, 6, 30)), key="p_date_box", label_visibility="collapsed")
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
        st.date_input("Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="gantt_date_top", label_visibility="collapsed")

    st.write("")
    # Filter Bar Card
    with st.container(border=True):
        gf1, gf2, gf3, gf4, gf_sp, gf_more = st.columns([1.15, 1.15, 1.15, 1.15, 0.8, 0.7])
        gantt_projects = ["All Projects", "Website Revamp", "IT Infrastructure Upgrade", "Sustainability Initiative", "Marketing Campaign", "Training Program"]
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

        # WBS dataset for Month, Week, and Day views matching the 3 Gantt screenshots
        month_day_rows = [
            # (id, indent, icon_type, icon_col, name, start, end, dur, status, pct, bar_left, bar_w, bar_col, bar_fill_w, milestone_txt, has_dep_arrow, row_bg)
            ("1", 0, "folder", "#00875A", "Website Revamp", "Apr 24, 2024", "Jun 30, 2024", "67 days", "On Track", "70%", 7.5, 83.0, "#00875A", 100, " end_diamond", False, "#ECFDF5"),
            ("1.1", 1, "folder", "#1473E6", "Planning", "Apr 24, 2024", "May 5, 2024", "12 days", "Completed", "100%", 7.5, 14.5, "#00875A", 100, "Project Kick-off", False, "#FFFFFF"),
            ("1.1.1", 2, "play", "#00875A", "Define project scope", "Apr 24, 2024", "Apr 26, 2024", "3 days", "Completed", "100%", 7.5, 4.5, "#10B981", 100, "", True, "#FFFFFF"),
            ("1.1.2", 2, "play", "#00875A", "Stakeholder alignment", "Apr 27, 2024", "May 5, 2024", "9 days", "Completed", "100%", 12.0, 12.5, "#10B981", 100, "", False, "#FFFFFF"),
            ("1.2", 1, "folder", "#00875A", "Design", "May 6, 2024", "May 20, 2024", "15 days", "Completed", "100%", 24.5, 20.0, "#00875A", 100, "Design Sign-off", False, "#FFFFFF"),
            ("1.2.1", 2, "play", "#00875A", "Create wireframes", "May 6, 2024", "May 12, 2024", "7 days", "Completed", "100%", 24.5, 11.0, "#10B981", 100, "", True, "#FFFFFF"),
            ("1.2.2", 2, "play", "#00875A", "UI/UX design", "May 13, 2024", "May 20, 2024", "8 days", "Completed", "100%", 35.5, 12.0, "#10B981", 100, "", False, "#FFFFFF"),
            ("1.3", 1, "folder", "#1473E6", "Development", "May 21, 2024", "May 31, 2024", "11 days", "In Progress", "60%", 41.5, 17.5, "#3B82F6", 60, "", False, "#F0F7FF"),
            ("1.3.1", 2, "play", "#64748B", "Frontend development", "May 21, 2024", "May 26, 2024", "6 days", "In Progress", "70%", 41.5, 9.5, "#60A5FA", 100, "", True, "#FFFFFF"),
            ("1.3.2", 2, "play", "#64748B", "Backend development", "May 27, 2024", "May 31, 2024", "5 days", "In Progress", "50%", 51.0, 9.5, "#60A5FA", 100, "", False, "#FFFFFF"),
            ("1.4", 1, "folder", "#64748B", "Testing", "Jun 1, 2024", "Jun 15, 2024", "15 days", "Not Started", "0%", 57.5, 19.0, "#8495B0", 100, "", True, "#FFFFFF"),
            ("1.4.1", 2, "play", "#64748B", "UAT testing", "Jun 1, 2024", "Jun 10, 2024", "10 days", "Not Started", "0%", 57.5, 12.0, "#8495B0", 100, "", False, "#FFFFFF"),
            ("1.4.2", 2, "play", "#64748B", "Bug fixes", "Jun 11, 2024", "Jun 15, 2024", "5 days", "Not Started", "0%", 69.5, 7.5, "#8495B0", 100, "", True, "#FFFFFF"),
            ("1.5", 1, "folder", "#7C3AED", "Go-Live", "Jun 16, 2024", "Jun 30, 2024", "15 days", "Not Started", "0%", 77.0, 17.0, "#8495B0", 100, "Go-Live", False, "#FFFFFF"),
            ("2", 0, "folder", "#00875A", "IT Infrastructure Upgrade", "Apr 15, 2024", "Jun 20, 2024", "67 days", "In Progress", "45%", 0.5, 69.0, "#93C5FD", 55, "", False, "#F0F7FF"),
            ("2.1", 1, "doc", "#00875A", "Requirements", "Apr 15, 2024", "Apr 30, 2024", "16 days", "Completed", "100%", 0.5, 17.5, "#00875A", 100, "", True, "#FFFFFF"),
            ("2.2", 1, "doc", "#7C3AED", "Implementation", "May 1, 2024", "Jun 10, 2024", "41 days", "In Progress", "40%", 18.5, 40.5, "#2563EB", 100, "", True, "#FFFFFF"),
            ("2.3", 1, "doc", "#7C3AED", "Testing", "Jun 11, 2024", "Jun 20, 2024", "10 days", "Not Started", "0%", 60.0, 16.0, "#8495B0", 100, "", False, "#FFFFFF"),
            ("3", 0, "folder", "#00875A", "Sustainability Initiative", "May 1, 2024", "Jun 30, 2024", "61 days", "At Risk", "30%", 0.5, 75.5, "#F59E0B", 100, "", False, "#FFFBEB"),
            ("3.1", 1, "doc", "#00875A", "Research & Analysis", "May 1, 2024", "May 20, 2024", "20 days", "Completed", "100%", 0.5, 29.0, "#00875A", 100, "", True, "#FFFFFF"),
            ("3.2", 1, "doc", "#64748B", "Strategy Development", "May 21, 2024", "Jun 10, 2024", "21 days", "In Progress", "40%", 30.8, 29.5, "#2563EB", 100, "", True, "#FFFFFF"),
            ("3.3", 1, "doc", "#64748B", "Implementation", "Jun 11, 2024", "Jun 30, 2024", "20 days", "Not Started", "0%", 61.0, 27.5, "#8495B0", 100, "Launch", False, "#FFFFFF"),
        ]

        week_rows = [
            ("1", 0, "folder", "#00875A", "Website Revamp", "Apr 24, 2024", "Jun 30, 2024", "67 days", "On Track", "70%", 7.5, 84.0, "#00875A", 100, "", False, "#ECFDF5"),
            ("1.1", 1, "folder", "#1473E6", "Planning", "Apr 24, 2024", "May 5, 2024", "12 days", "Completed", "100%", 7.5, 14.5, "#10B981", 100, "", True, "#FFFFFF"),
            ("1.2", 1, "folder", "#00875A", "Design", "May 6, 2024", "May 20, 2024", "15 days", "Completed", "100%", 24.0, 17.0, "#10B981", 100, "", True, "#FFFFFF"),
            ("1.3", 1, "folder", "#1473E6", "Development", "May 21, 2024", "May 31, 2024", "11 days", "In Progress", "60%", 43.0, 16.5, "#2563EB", 100, "", True, "#FFFFFF"),
            ("1.4", 1, "folder", "#7C3AED", "Testing", "Jun 1, 2024", "Jun 15, 2024", "15 days", "Not Started", "0%", 60.0, 17.5, "#8495B0", 100, "", True, "#FFFFFF"),
            ("1.5", 1, "folder", "#475569", "Go-Live", "Jun 16, 2024", "Jun 30, 2024", "15 days", "Not Started", "0%", 78.0, 15.0, "#8495B0", 100, "Go-Live", False, "#FFFFFF"),
            ("2", 0, "folder", "#00875A", "IT Infrastructure Upgrade", "Apr 15, 2024", "Jun 20, 2024", "67 days", "In Progress", "45%", 6.5, 74.5, "#93C5FD", 46, "", False, "#F0F7FF"),
            ("2.1", 1, "doc", "#00875A", "Requirements", "Apr 15, 2024", "Apr 30, 2024", "16 days", "Completed", "100%", 6.5, 17.5, "#10B981", 100, "", True, "#FFFFFF"),
            ("2.2", 1, "doc", "#1473E6", "Implementation", "May 1, 2024", "Jun 10, 2024", "41 days", "In Progress", "40%", 24.8, 43.0, "#3B82F6", 100, "", True, "#FFFFFF"),
            ("2.3", 1, "doc", "#7C3AED", "Testing", "Jun 11, 2024", "Jun 20, 2024", "10 days", "Not Started", "0%", 68.2, 16.5, "#8495B0", 100, "", False, "#FFFFFF"),
            ("3", 0, "folder", "#00875A", "Sustainability Initiative", "May 1, 2024", "Jun 30, 2024", "61 days", "At Risk", "30%", 6.5, 78.5, "#F59E0B", 100, "", False, "#FFFBEB"),
            ("3.1", 1, "doc", "#00875A", "Research & Analysis", "May 1, 2024", "May 20, 2024", "20 days", "Completed", "100%", 6.5, 25.0, "#00875A", 100, "", True, "#FFFFFF"),
            ("3.2", 1, "doc", "#3B82F6", "Strategy Development", "May 21, 2024", "Jun 10, 2024", "21 days", "In Progress", "40%", 33.0, 25.5, "#3B82F6", 100, "", True, "#FFFFFF"),
            ("3.3", 1, "doc", "#7C3AED", "Implementation", "Jun 11, 2024", "Jun 30, 2024", "20 days", "Not Started", "0%", 59.5, 32.0, "#8495B0", 100, "Launch", False, "#FFFFFF"),
            ("4", 0, "folder", "#00875A", "Marketing Campaign", "May 6, 2024", "Jun 25, 2024", "50 days", "On Track", "25%", 19.8, 62.0, "#00875A", 100, "", False, "#ECFDF5"),
            ("4.1", 1, "doc", "#1473E6", "Content Creation", "May 6, 2024", "May 20, 2024", "15 days", "In Progress", "50%", 19.8, 20.0, "#2563EB", 100, "", True, "#FFFFFF"),
            ("4.2", 1, "doc", "#7C3AED", "Campaign Setup", "May 21, 2024", "Jun 10, 2024", "20 days", "Not Started", "0%", 40.5, 21.0, "#8495B0", 100, "", True, "#FFFFFF"),
            ("4.3", 1, "doc", "#475569", "Launch Campaign", "Jun 11, 2024", "Jun 25, 2024", "15 days", "Not Started", "0%", 62.0, 22.0, "#8495B0", 100, "Launch", False, "#FFFFFF"),
            ("5", 0, "folder", "#00875A", "Training Program", "May 13, 2024", "Jun 28, 2024", "47 days", "On Track", "20%", 29.5, 56.0, "#00875A", 100, "", False, "#ECFDF5"),
            ("5.1", 1, "doc", "#00875A", "Training Materials", "May 13, 2024", "May 31, 2024", "18 days", "In Progress", "40%", 29.5, 23.8, "#2563EB", 100, "", True, "#FFFFFF"),
            ("5.2", 1, "doc", "#475569", "Pilot Training", "Jun 3, 2024", "Jun 14, 2024", "12 days", "Not Started", "0%", 55.0, 14.0, "#8495B0", 100, "", True, "#FFFFFF"),
            ("5.3", 1, "doc", "#475569", "Full Rollout", "Jun 17, 2024", "Jun 28, 2024", "11 days", "Not Started", "0%", 70.8, 18.0, "#8495B0", 100, "Complete", False, "#FFFFFF"),
        ]

        day_rows = [
            ("1", 0, "folder", "#00875A", "Website Revamp", "Apr 24, 2024", "Jun 30, 2024", "67 days", "On Track", "70%", 0.0, 100.0, "#00875A", 100, "", False, "#ECFDF5", ""),
            ("1.1", 1, "folder", "#00875A", "Planning", "Apr 24, 2024", "May 5, 2024", "12 days", "Completed", "100%", 0.0, 9.5, "#A7F3D0", 100, "Project Kick-off", False, "#FFFFFF", ""),
            ("1.2", 1, "folder", "#00875A", "Design", "May 6, 2024", "May 20, 2024", "15 days", "Completed", "100%", 0.0, 0.0, "#00875A", 100, "", False, "#FFFFFF", ""),
            ("1.2.1", 2, "play", "#00875A", "Create wireframes", "May 6, 2024", "May 12, 2024", "7 days", "Completed", "100%", 3.5, 24.5, "#00875A", 100, "", True, "#FFFFFF", "Create wireframes"),
            ("1.2.2", 2, "play", "#00875A", "UI/UX design", "May 13, 2024", "May 20, 2024", "8 days", "Completed", "100%", 29.5, 17.5, "#10B981", 100, "", False, "#FFFFFF", "UI/UX design"),
            ("1.3", 1, "folder", "#1473E6", "Development", "May 21, 2024", "May 31, 2024", "11 days", "In Progress", "60%", 25.8, 20.5, "#60A5FA", 45, "", False, "#F0F7FF", "Frontend development"),
            ("1.3.1", 2, "play", "#0F766E", "Frontend development", "May 21, 2024", "May 26, 2024", "6 days", "In Progress", "70%", 48.5, 15.0, "#60A5FA", 100, "", True, "#FFFFFF", "Backend development"),
            ("1.3.2", 2, "play", "#0F766E", "Backend development", "May 27, 2024", "May 31, 2024", "5 days", "In Progress", "50%", 66.8, 16.0, "#8495B0", 100, "", False, "#FFFFFF", ""),
            ("1.4", 1, "folder", "#475569", "Testing", "Jun 1, 2024", "Jun 15, 2024", "15 days", "Not Started", "0%", 58.8, 12.5, "#8495B0", 100, "", True, "#FFFFFF", "UAT testing"),
            ("1.5", 1, "folder", "#7C3AED", "Go-Live", "Jun 16, 2024", "Jun 30, 2024", "15 days", "Not Started", "0%", 73.0, 14.0, "#8495B0", 100, "Go-Live", False, "#FFFFFF", ""),
            ("2", 0, "folder", "#00875A", "IT Infrastructure Upgrade", "Apr 15, 2024", "Jun 20, 2024", "67 days", "In Progress", "45%", 0.0, 71.5, "#60A5FA", 49, "", False, "#F0F7FF", ""),
            ("2.1", 1, "doc", "#00875A", "Requirements", "Apr 15, 2024", "Apr 30, 2024", "16 days", "Completed", "100%", 0.0, 16.0, "#10B981", 100, "", True, "#FFFFFF", "Requirements"),
            ("2.2", 1, "doc", "#1473E6", "Implementation", "May 1, 2024", "Jun 10, 2024", "41 days", "In Progress", "40%", 18.0, 33.2, "#3B82F6", 100, "", True, "#FFFFFF", "Implementation"),
            ("2.3", 1, "doc", "#7C3AED", "Testing", "Jun 11, 2024", "Jun 20, 2024", "10 days", "Not Started", "0%", 53.0, 23.5, "#8495B0", 100, "", False, "#FFFFFF", "Testing"),
            ("3", 0, "folder", "#00875A", "Sustainability Initiative", "May 1, 2024", "Jun 30, 2024", "61 days", "At Risk", "30%", 0.0, 88.5, "#F59E0B", 100, "", False, "#FFFBEB", ""),
            ("3.1", 1, "doc", "#00875A", "Research & Analysis", "May 1, 2024", "May 20, 2024", "20 days", "Completed", "100%", 0.0, 20.5, "#00875A", 100, "", True, "#FFFFFF", "Research & Analysis"),
            ("3.2", 1, "doc", "#1473E6", "Strategy Development", "May 21, 2024", "Jun 10, 2024", "21 days", "In Progress", "40%", 22.5, 19.5, "#3B82F6", 100, "", True, "#FFFFFF", "Strategy Development"),
            ("3.3", 1, "doc", "#7C3AED", "Implementation", "Jun 11, 2024", "Jun 30, 2024", "20 days", "Not Started", "0%", 43.8, 24.5, "#8495B0", 100, "Complete", False, "#FFFFFF", "Implementation"),
        ]

        active_rows = week_rows if scale == "Week" else (day_rows if scale == "Day" else month_day_rows)

        # Apply Top Search & Dropdown Filters
        proj_prefix_map = {
            "Website Revamp": "1",
            "IT Infrastructure Upgrade": "2",
            "Sustainability Initiative": "3",
            "Marketing Campaign": "4",
            "Training Program": "5",
        }
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

        # Build Timeline Header HTML according to scale
        if scale == "Month":
            tl_top_hdr = """
            <div style="display:grid;grid-template-columns:2fr 5fr 5fr;border-bottom:1px solid #E2E8F0;background:#F8FAFC;font-size:11.5px;font-weight:700;color:#1E293B;text-align:center">
                <div style="padding:6px 0;border-right:1px solid #E2E8F0">Apr 2024</div>
                <div style="padding:6px 0;border-right:1px solid #E2E8F0;position:relative">
                    May 2024
                    <span style="position:absolute;left:10%;top:-10px;background:#065F46;color:white;font-size:10px;font-weight:700;padding:2px 8px;border-radius:4px;z-index:5">Today</span>
                </div>
                <div style="padding:6px 0">Jun 2024</div>
            </div>
            <div style="display:grid;grid-template-columns:repeat(12,1fr);background:#F8FAFC;font-size:10.5px;font-weight:600;color:#475569;text-align:center">
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">Apr 21</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">Apr 28</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">May 5</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">May 12</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7;color:#1D4ED8">May 19</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">May 26</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">May 26</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7;color:#1D4ED8">Jun 2</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7;color:#1D4ED8">Jun 9</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7;color:#1D4ED8">Jun 16</div>
                <div style="padding:5px 0;border-right:1px solid #EDF2F7">Jun 23</div>
                <div style="padding:5px 0">Jun 30</div>
            </div>
            """
            today_line_html = '<div style="position:absolute;top:0;bottom:0;left:21%;width:1.5px;background:#10B981;z-index:4;pointer-events:none"></div>'
        elif scale == "Week":
            weeks_meta = [
                ("Apr 22 – Apr 28", "Week 17", False),
                ("Apr 29 – May 5", "Week 18", False),
                ("May 6 – May 12", "Week 19", True),
                ("May 13 – May 19", "Week 20", False),
                ("May 20 – May 26", "Week 21", False),
                ("May 27 – Jun 2", "Week 22", False),
                ("Jun 3 – Jun 9", "Week 23", False),
                ("Jun 10 – Jun 16", "Week 24", False),
                ("Jun 17 – Jun 23", "Week 25", False),
                ("Jun 24 – Jun 30", "Week 26", False),
            ]
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
                <div style="font-size:12px;font-weight:700;color:#0F172A">‹ &nbsp;&nbsp;&nbsp; Mon, May 6, 2024 &nbsp;&nbsp;&nbsp; ›</div>
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
            parent_t = st.selectbox("Parent Task", ["None (Top Level Task)", "1.1 Define project scope", "1.2 Stakeholder alignment", "1.3 Create wireframes"], key="add_t_parent")

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
            t_start = st.date_input("Start Date *", date(2024, 4, 24), key="add_t_start")
        with r4_b:
            t_end = st.date_input("End Date *", date(2024, 4, 30), key="add_t_end")
        with r4_c:
            st.number_input("Duration (Days)", min_value=1, max_value=365, value=7, key="add_t_dur")

        r5_a, r5_b, r5_c = st.columns([1.1, 1.1, 1])
        with r5_a:
            assignees_pool = ["Dana Scott", "Emily Rodriguez", "James Miller", "Sarah Kim", "Alice Morgan", "Ryan Cruz", "Michael Tan"]
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

        st.markdown(
            clean_html(
                """
                <div style="border:1px solid #EDF2F7;border-radius:8px;padding:8px 12px;margin-top:6px;margin-bottom:12px;font-size:12.5px;color:#1E293B">
                    <div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #F1F5F9">
                        <span>☐ &nbsp; Review requirements document</span><span style="color:#64748B">🗑</span>
                    </div>
                    <div style="display:flex;justify-content:space-between;padding:5px 0">
                        <span>☐ &nbsp; Finalize deliverables</span><span style="color:#64748B">🗑</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

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
                sql(
                    "INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (
                        t_name.strip() or "New Project Task",
                        pmap.get(proj, 1),
                        "IT",
                        clean_stat,
                        clean_prio,
                        t_start.isoformat(),
                        t_end.isoformat(),
                        t_ass,
                        int(t_pct),
                        "" if t_dep == "Select dependent tasks" else t_dep,
                        t_notes or t_desc,
                    ),
                )
                log_event("Created", "Task", None, t_name.strip() or "New Project Task")
                if not create_another:
                    st.session_state.pop("task_modal", None)
                st.rerun()

    @st.dialog("View Task", width="large")
    def view_task_modal(task_rec: dict) -> None:
        tname = task_rec.get("name", "Create wireframes")
        tproj = task_rec.get("project", "Website Revamp")
        tstart = task_rec.get("start_date", "May 6, 2024")
        tend = task_rec.get("end_date", "May 12, 2024")
        tdur = task_rec.get("duration", "7 days")
        tpct = int(task_rec.get("pct", 100))
        tstat = task_rec.get("status", "Completed")
        tprio = task_rec.get("priority", "High")
        tass = task_rec.get("assignee", "James Miller")

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
                            Create low and high-fidelity wireframes for the new website based on the approved requirements and UX guidelines. Include key user flows, page layouts and navigation structure for review.
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
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">☰ &nbsp;1.2 Stakeholder alignment</div>
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
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">☑ &nbsp;Checklist / Subtasks (3)</div>
                        <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
                            <div style="flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden"><div style="width:100%;height:100%;background:#00875A"></div></div>
                            <span style="font-size:11.5px;color:#475569;font-weight:600">3 of 3 completed</span>
                        </div>
                        <div style="font-size:12.5px;line-height:2.1">
                            <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span>✅ &nbsp;Create homepage wireframe</span><span style="color:#64748B">May 8, 2024</span></div>
                            <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span>✅ &nbsp;Create product pages wireframe</span><span style="color:#64748B">May 10, 2024</span></div>
                            <div style="display:flex;justify-content:space-between"><span>✅ &nbsp;Review with stakeholders</span><span style="color:#64748B">May 12, 2024</span></div>
                        </div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">📎 &nbsp;Attachments (2)</div>
                        <div style="border:1px solid #EDF2F7;border-radius:8px;padding:8px 10px;display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
                            <div style="display:flex;align-items:center;gap:10px">
                                <span style="background:#EF4444;color:white;font-size:10px;font-weight:700;padding:6px 6px;border-radius:6px">PDF</span>
                                <div><div style="font-size:12.5px;font-weight:700">Wireframe_Overview.pdf</div><div style="font-size:11px;color:#64748B">2.4 MB • Apr 30, 2024</div></div>
                            </div>
                            <span style="color:#475569;font-weight:700">⬇ &nbsp;⋮</span>
                        </div>
                        <div style="border:1px solid #EDF2F7;border-radius:8px;padding:8px 10px;display:flex;align-items:center;justify-content:space-between">
                            <div style="display:flex;align-items:center;gap:10px">
                                <span style="background:#1473E6;color:white;font-size:10px;font-weight:700;padding:6px 6px;border-radius:6px">IMG</span>
                                <div><div style="font-size:12.5px;font-weight:700">Homepage_Wireframe.png</div><div style="font-size:11px;color:#64748B">1.1 MB • May 1, 2024</div></div>
                            </div>
                            <span style="color:#475569;font-weight:700">⬇ &nbsp;⋮</span>
                        </div>
                    </div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white;margin-bottom:14px">
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">💬 &nbsp;Recent Activity / Notes</div>
                    <div style="font-size:12.5px;line-height:2.2">
                        <div style="display:grid;grid-template-columns:150px 155px 1fr;gap:8px;border-bottom:1px solid #F1F5F9;align-items:center">
                            <span>{avatar_html('James Miller')}</span>
                            <span style="color:#64748B;font-size:11.5px">May 12, 2024 3:45 PM</span>
                            <span>Wireframes completed and approved by stakeholders.</span>
                        </div>
                        <div style="display:grid;grid-template-columns:150px 155px 1fr;gap:8px;align-items:center;padding-top:4px">
                            <span>{avatar_html('Emily Rodriguez')}</span>
                            <span style="color:#64748B;font-size:11.5px">May 10, 2024 10:20 AM</span>
                            <span>Initial wireframes shared for review. Feedback incorporated.</span>
                        </div>
                    </div>
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
                st.session_state.task_modal = ("edit", task_rec)
                st.rerun()

    @st.dialog("Edit Task", width="large")
    def edit_task_modal(task_rec: dict) -> None:
        tname = task_rec.get("name", "Create wireframes")
        tproj = task_rec.get("project", "Website Revamp")
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
                value="Create low and high-fidelity wireframes for the new website based on the approved requirements and UX guidelines. Include key user flows, page layouts and navigation structure for review.",
                height=92,
                key="ed_t_desc",
            )
        with e2_r:
            ass_list = ["James Miller (UX Designer)", "Dana Scott", "Emily Rodriguez", "Sarah Kim", "Alice Morgan", "Ryan Cruz", "Michael Tan"]
            ed_ass = st.selectbox("👤 Assignee *", ass_list, index=0, key="ed_t_ass")

        e3_a, e3_b, e3_c, e3_d = st.columns(4)
        with e3_a:
            ed_start = st.date_input("Start Date *", date(2024, 5, 6), key="ed_t_start")
        with e3_b:
            ed_end = st.date_input("End Date *", date(2024, 5, 12), key="ed_t_end")
        with e3_c:
            st.number_input("Duration (Days) *", min_value=1, max_value=365, value=7, key="ed_t_dur")
        with e3_d:
            ed_pct = st.number_input("% Complete *", min_value=0, max_value=100, value=int(task_rec.get("pct", 100)), key="ed_t_pct")

        e4_a, e4_b, e4_c, e4_d = st.columns(4)
        with e4_a:
            ed_prio = st.selectbox("Priority *", ["⚑ High", "⚑ Medium", "⚑ Low"], index=0, key="ed_t_prio")
        with e4_b:
            ed_stat = st.selectbox("Status *", ["🟢 Completed", "🔵 In Progress", "⚪ Not Started", "🟠 At Risk"], index=0, key="ed_t_stat")
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
                ed_notes = st.text_area("Progress Notes", value="Wireframes completed and approved by stakeholders.", height=68, key="ed_p_notes")

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
                log_event("Updated", "Task", None, ed_name)
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
        st.date_input("Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="tl_d_top", label_visibility="collapsed")
    with th_b:
        if st.button("＋ Add Task", type="primary", use_container_width=True, key="tl_btn_add_top"):
            st.session_state.task_modal = ("add", None)
            st.rerun()

    st.write("")
    # --- FILTER ROW ---
    tf1, tf2, tf3, tf4, tf5, tf6 = st.columns([1.05, 1.0, 1.0, 1.05, 1.35, 0.75])
    with tf1:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Project</div>", unsafe_allow_html=True)
        tl_proj_f = st.selectbox("Project", ["All Projects", "Website Revamp", "IT Infrastructure Upgrade", "Sustainability Initiative", "Marketing Campaign"], key="tl_f_proj", label_visibility="collapsed")
    with tf2:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Status</div>", unsafe_allow_html=True)
        tl_stat_f = st.selectbox("Status", ["All Statuses", "Completed", "In Progress", "Not Started", "At Risk", "Overdue"], key="tl_f_stat", label_visibility="collapsed")
    with tf3:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Priority</div>", unsafe_allow_html=True)
        tl_prio_f = st.selectbox("Priority", ["All Priorities", "High", "Medium", "Low"], key="tl_f_prio", label_visibility="collapsed")
    with tf4:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Assignee</div>", unsafe_allow_html=True)
        tl_ass_f = st.selectbox("Assignee", ["All Assignees", "Dana Scott", "Emily Rodriguez", "James Miller", "Sarah Kim", "Alice Morgan", "Ryan Cruz", "Michael Tan"], key="tl_f_ass", label_visibility="collapsed")
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
    kpi_specs = [
        (tk1, "#1473E6", "#FFFFFF", "▤", "Total Tasks", "#475569", "48", ""),
        (tk2, "#00875A", "#FFFFFF", "✓", "Completed", "#475569", "18", "(38%)"),
        (tk3, "#1473E6", "#FFFFFF", "❂", "In Progress", "#475569", "16", "(33%)"),
        (tk4, "#64748B", "#FFFFFF", "●", "Not Started", "#475569", "8", "(17%)"),
        (tk5, "#F28C00", "#FFFFFF", "!", "At Risk", "#D97706", "4", "(8%)"),
        (tk6, "#EF4444", "#FFFFFF", "!", "Overdue", "#DC2626", "2", "(4%)"),
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
    task_groups = [
        {
            "group_title": "Website Revamp (8)",
            "project": "Website Revamp",
            "icon_bg": "#00875A",
            "start": "Apr 24, 2024",
            "end": "Jun 30, 2024",
            "dur": "67 days",
            "pct": 70,
            "bar_col": "#00875A",
            "row_bg": "#F0FDF4",
            "tasks": [
                {"wbs": "1.1", "ico": "doc", "ic_col": "#00875A", "name": "Define project scope", "project": "Website Revamp", "start_date": "Apr 24, 2024", "end_date": "Apr 26, 2024", "duration": "3 days", "status": "Completed", "priority": "Medium", "assignee": "Dana Scott", "pct": 100, "bar_col": "#00875A", "tag": "Planning", "tag_bg": "#DBEAFE", "tag_fg": "#1D4ED8"},
                {"wbs": "1.2", "ico": "play", "ic_col": "#00875A", "name": "Stakeholder alignment", "project": "Website Revamp", "start_date": "Apr 27, 2024", "end_date": "May 5, 2024", "duration": "9 days", "status": "Completed", "priority": "Medium", "assignee": "Emily Rodriguez", "pct": 100, "bar_col": "#00875A", "tag": "Planning", "tag_bg": "#DBEAFE", "tag_fg": "#1D4ED8"},
                {"wbs": "1.3", "ico": "play", "ic_col": "#00875A", "name": "Create wireframes", "project": "Website Revamp", "start_date": "May 6, 2024", "end_date": "May 12, 2024", "duration": "7 days", "status": "Completed", "priority": "High", "assignee": "James Miller", "pct": 100, "bar_col": "#00875A", "tag": "Design", "tag_bg": "#F3E8FF", "tag_fg": "#7C3AED"},
                {"wbs": "1.4", "ico": "play", "ic_col": "#00875A", "name": "UI/UX design", "project": "Website Revamp", "start_date": "May 13, 2024", "end_date": "May 20, 2024", "duration": "8 days", "status": "Completed", "priority": "High", "assignee": "Sarah Kim", "pct": 100, "bar_col": "#00875A", "tag": "Design", "tag_bg": "#F3E8FF", "tag_fg": "#7C3AED"},
                {"wbs": "1.5", "ico": "play", "ic_col": "#00875A", "name": "Frontend development", "project": "Website Revamp", "start_date": "May 21, 2024", "end_date": "May 26, 2024", "duration": "6 days", "status": "In Progress", "priority": "High", "assignee": "Alice Morgan", "pct": 70, "bar_col": "#1473E6", "tag": "Development", "tag_bg": "#DCFCE7", "tag_fg": "#15803D"},
                {"wbs": "1.6", "ico": "play", "ic_col": "#0F766E", "name": "Backend development", "project": "Website Revamp", "start_date": "May 27, 2024", "end_date": "May 31, 2024", "duration": "5 days", "status": "In Progress", "priority": "High", "assignee": "Ryan Cruz", "pct": 50, "bar_col": "#1473E6", "tag": "Development", "tag_bg": "#DCFCE7", "tag_fg": "#15803D"},
                {"wbs": "1.7", "ico": "doc", "ic_col": "#475569", "name": "UAT testing", "project": "Website Revamp", "start_date": "Jun 1, 2024", "end_date": "Jun 15, 2024", "duration": "15 days", "status": "Not Started", "priority": "Medium", "assignee": "Michael Tan", "pct": 0, "bar_col": "#CBD5E1", "tag": "Testing", "tag_bg": "#FEE2E2", "tag_fg": "#B91C1C"},
                {"wbs": "1.8", "ico": "doc", "ic_col": "#7C3AED", "name": "Go-Live", "project": "Website Revamp", "start_date": "Jun 16, 2024", "end_date": "Jun 30, 2024", "duration": "15 days", "status": "Not Started", "priority": "High", "assignee": "Alice Morgan", "pct": 0, "bar_col": "#CBD5E1", "tag": "Launch", "tag_bg": "#EDE9FE", "tag_fg": "#6D28D9"},
            ],
        },
        {
            "group_title": "IT Infrastructure Upgrade (5)",
            "project": "IT Infrastructure Upgrade",
            "icon_bg": "#1473E6",
            "start": "Apr 15, 2024",
            "end": "Jun 20, 2024",
            "dur": "67 days",
            "pct": 45,
            "bar_col": "#1473E6",
            "row_bg": "#EFF6FF",
            "tasks": [
                {"wbs": "2.1", "ico": "doc", "ic_col": "#0F766E", "name": "Requirements", "project": "IT Infrastructure..", "start_date": "Apr 15, 2024", "end_date": "Apr 30, 2024", "duration": "16 days", "status": "Completed", "priority": "High", "assignee": "Dana Scott", "pct": 100, "bar_col": "#00875A", "tag": "Requirements", "tag_bg": "#E2E8F0", "tag_fg": "#334155"},
                {"wbs": "2.2", "ico": "doc", "ic_col": "#1473E6", "name": "Implementation", "project": "IT Infrastructure..", "start_date": "May 1, 2024", "end_date": "Jun 10, 2024", "duration": "41 days", "status": "In Progress", "priority": "High", "assignee": "Ryan Cruz", "pct": 40, "bar_col": "#1473E6", "tag": "Implementation", "tag_bg": "#E2E8F0", "tag_fg": "#334155"},
                {"wbs": "2.3", "ico": "doc", "ic_col": "#7C3AED", "name": "Testing", "project": "IT Infrastructure..", "start_date": "Jun 11, 2024", "end_date": "Jun 20, 2024", "duration": "10 days", "status": "Not Started", "priority": "Medium", "assignee": "Michael Tan", "pct": 0, "bar_col": "#CBD5E1", "tag": "Testing", "tag_bg": "#FEE2E2", "tag_fg": "#B91C1C"},
            ],
        },
        {
            "group_title": "Sustainability Initiative (4)",
            "project": "Sustainability Initiative",
            "icon_bg": "#7C3AED",
            "start": "May 1, 2024",
            "end": "Jun 30, 2024",
            "dur": "61 days",
            "pct": 30,
            "bar_col": "#F28C00",
            "row_bg": "#F5F3FF",
            "tasks": [
                {"wbs": "3.1", "ico": "doc", "ic_col": "#00875A", "name": "Research & Analysis", "project": "Sustainability In..", "start_date": "May 1, 2024", "end_date": "May 20, 2024", "duration": "20 days", "status": "Completed", "priority": "Medium", "assignee": "Emily Rodriguez", "pct": 100, "bar_col": "#00875A", "tag": "Research", "tag_bg": "#DCFCE7", "tag_fg": "#15803D"},
                {"wbs": "3.2", "ico": "doc", "ic_col": "#1473E6", "name": "Strategy Development", "project": "Sustainability In..", "start_date": "May 21, 2024", "end_date": "Jun 10, 2024", "duration": "21 days", "status": "In Progress", "priority": "Medium", "assignee": "Sarah Kim", "pct": 40, "bar_col": "#1473E6", "tag": "Strategy", "tag_bg": "#DBEAFE", "tag_fg": "#1D4ED8"},
                {"wbs": "3.3", "ico": "doc", "ic_col": "#7C3AED", "name": "Implementation", "project": "Sustainability In..", "start_date": "Jun 11, 2024", "end_date": "Jun 30, 2024", "duration": "20 days", "status": "Not Started", "priority": "Low", "assignee": "James Miller", "pct": 0, "bar_col": "#CBD5E1", "tag": "Implementation", "tag_bg": "#F3E8FF", "tag_fg": "#7C3AED"},
            ],
        },
        {
            "group_title": "Marketing Campaign (3)",
            "project": "Marketing Campaign",
            "icon_bg": "#F28C00",
            "start": "May 6, 2024",
            "end": "Jun 25, 2024",
            "dur": "50 days",
            "pct": 25,
            "bar_col": "#00875A",
            "row_bg": "#FFFBEB",
            "tasks": [
                {"wbs": "4.1", "ico": "doc", "ic_col": "#00875A", "name": "Content Creation", "project": "Marketing Campaig", "start_date": "May 6, 2024", "end_date": "May 20, 2024", "duration": "15 days", "status": "In Progress", "priority": "Medium", "assignee": "Alice Morgan", "pct": 50, "bar_col": "#1473E6", "tag": "Content", "tag_bg": "#DBEAFE", "tag_fg": "#1D4ED8"},
                {"wbs": "4.2", "ico": "doc", "ic_col": "#1473E6", "name": "Campaign Setup", "project": "Marketing Campaig", "start_date": "May 21, 2024", "end_date": "Jun 10, 2024", "duration": "20 days", "status": "Not Started", "priority": "Low", "assignee": "Ryan Cruz", "pct": 0, "bar_col": "#CBD5E1", "tag": "Setup", "tag_bg": "#EDE9FE", "tag_fg": "#6D28D9"},
                {"wbs": "4.3", "ico": "doc", "ic_col": "#475569", "name": "Launch Campaign", "project": "Marketing Campaig", "start_date": "Jun 11, 2024", "end_date": "Jun 25, 2024", "duration": "15 days", "status": "Not Started", "priority": "High", "assignee": "Michael Tan", "pct": 0, "bar_col": "#CBD5E1", "tag": "Launch", "tag_bg": "#EDE9FE", "tag_fg": "#6D28D9"},
            ],
        },
    ]

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
                        if st.button("👁", key=f"tl_v_{wbs_id}", help="View Task"):
                            st.session_state.task_modal = ("view", t_item)
                            st.rerun()
                    with b_e:
                        if st.button("❐", key=f"tl_e_{wbs_id}", help="Edit Task"):
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
                st.date_input("Availability Start Date *", date(2024, 4, 24), key="ar_av_start")
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
        rname = res_rec.get("name", "James Miller")
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
        with tb1:
            st.markdown(
                clean_html(
                    """
                    <div style="display:flex;align-items:center;gap:10px;padding-top:4px">
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 10px;background:white;font-weight:700">‹</span>
                        <span style="font-size:15px;font-weight:800;color:#0F172A">May 2024</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 10px;background:white;font-weight:700">›</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
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
            rc_weeks = [
                [("28", False, []), ("29", False, []), ("30", False, []), ("1", True, [("Website Revamp", "(50%)", "g"), ("Design Review", "(2 hrs)", "p")]), ("2", True, [("Website Revamp", "(50%)", "g")]), ("3", True, [("Website Revamp", "(50%)", "g")]), ("4", True, [])],
                [("5", True, []), ("6", True, [("Marketing Campaign", "(30%)", "g")]), ("7", True, [("Marketing Campaign", "(30%)", "g")]), ("8", True, [("Marketing Campaign", "(30%)", "g")]), ("9", True, [("Website Revamp", "(50%)", "g")]), ("10", True, [("Website Revamp", "(50%)", "g")]), ("11", True, [])],
                [("12", True, []), ("13", True, [("IT Infrastructure", "(20%)", "g")]), ("14", True, [("IT Infrastructure", "(20%)", "g")]), ("15", True, [("IT Infrastructure", "(20%)", "g"), ("Team Meeting", "(1 hr)", "p")]), ("16", True, [("IT Infrastructure", "(20%)", "g")]), ("17", True, [("IT Infrastructure", "(20%)", "g")]), ("18", True, [])],
                [("19", True, []), ("20", True, [("Annual Leave", "(Full Day)", "r")]), ("21", True, [("Annual Leave", "(Full Day)", "r")]), ("22", True, [("Annual Leave", "(Full Day)", "r")]), ("23", True, []), ("24", True, [("Website Revamp", "(50%)", "g")]), ("25", True, [])],
                [("26", True, []), ("27", True, [("Website Revamp", "(50%)", "g")]), ("28", True, [("Website Revamp", "(50%)", "g")]), ("29", True, [("Website Revamp", "(50%)", "g")]), ("30", True, [("Website Revamp", "(50%)", "g")]), ("31", True, [("Website Revamp", "(50%)", "g")]), ("1", False, [])],
            ]
            pill_styles = {
                "g": ("#DCFCE7", "#15803D"),
                "p": ("#F3E8FF", "#7C3AED"),
                "r": ("#FEE2E2", "#B91C1C"),
            }
            grid_rows_html = ""
            for wk in rc_weeks:
                cells_html = ""
                for d_str, in_m, ev_list in wk:
                    is_sel = in_m and d_str == "9"
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
        rname = res_rec.get("name", "James Miller")
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
        rname = res_rec.get("name", "James Miller")
        rrole = res_rec.get("role", "UX Designer")
        remail = res_rec.get("email", "james.miller@hpe.com")
        rini = res_rec.get("ini", "JM")
        rcolor = res_rec.get("av_col", "#0284C7")
        rskills = res_rec.get("skills", ["UI/UX", "Figma"])
        rproj = res_rec.get("project", "Website Revamp")
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
            st.date_input("Availability Start Date *", date(2024, 4, 24), key="sub_av_st")
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
        st.date_input("Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="res_d_top", label_visibility="collapsed")
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
        res_role_f = st.selectbox("Role", ["All Roles", "UX Designer", "Business Analyst", "Frontend Developer", "Project Manager", "Backend Developer", "QA Engineer", "Content Specialist", "DevOps Engineer", "IT Support", "Data Analyst"], key="res_f_role", label_visibility="collapsed")
    with rf_sk:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Skill</div>", unsafe_allow_html=True)
        res_skill_f = st.selectbox("Skill", ["All Skills", "UI/UX", "Figma", "Analysis", "React", "Planning", "Python", "Testing", "Content", "Azure", "Infrastructure", "Power BI"], key="res_f_skill", label_visibility="collapsed")
    with rf_st:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Status</div>", unsafe_allow_html=True)
        res_stat_f = st.selectbox("Status", ["All Statuses", "Allocated", "Available", "Overallocated", "On Leave"], key="res_f_stat", label_visibility="collapsed")
    with rf_pr:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Project</div>", unsafe_allow_html=True)
        res_proj_f = st.selectbox("Project", ["All Projects", "Website Revamp", "Sustainability Initiative", "Marketing Campaign", "IT Infrastructure Upgrade", "Internal Operations"], key="res_f_proj", label_visibility="collapsed")
    with rf_mr:
        st.markdown("<div style='font-size:11.5px;margin-bottom:3px'>&nbsp;</div>", unsafe_allow_html=True)
        if st.button("▽ More Filters", key="res_btn_more", use_container_width=True):
            for k, v in [("res_f_role", "All Roles"), ("res_f_skill", "All Skills"), ("res_f_stat", "All Statuses"), ("res_f_proj", "All Projects"), ("res_q_top", ""), ("res_q_sub", "")]:
                st.session_state[k] = v
            st.rerun()

    st.write("")
    # --- RESOURCES DIRECTORY TABLE (Exact 10 Rows from Image 20 & 25) ---
    res_directory = [
        {"id": 1, "ini": "JM", "av_col": "#0284C7", "name": "James Miller", "email": "james.miller@hpe.com", "role": "UX Designer", "skills": ["UI/UX", "Figma"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F3E8FF", "#7C3AED")], "project": "Website Revamp", "alloc": 100, "avail": "0%", "status": "Allocated"},
        {"id": 2, "ini": "EC", "av_col": "#0284C7", "name": "Emily Rodriguez", "email": "emily.rodriguez@hpe.com", "role": "Business Analyst", "skills": ["Analysis", "Requirements"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "Sustainability Initiative", "alloc": 80, "avail": "20%", "status": "Allocated"},
        {"id": 3, "ini": "SL", "av_col": "#0369A1", "name": "Sarah Kim", "email": "sarah.kim@hpe.com", "role": "Frontend Developer", "skills": ["React", "JavaScript"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "Website Revamp", "alloc": 100, "avail": "0%", "status": "Allocated"},
        {"id": 4, "ini": "AL", "av_col": "#0284C7", "name": "Alice Morgan", "email": "alice.morgan@hpe.com", "role": "Project Manager", "skills": ["Planning", "Stakeholder"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "Marketing Campaign", "alloc": 50, "avail": "50%", "status": "Available"},
        {"id": 5, "ini": "RJ", "av_col": "#0369A1", "name": "Ryan Cruz", "email": "ryan.cruz@hpe.com", "role": "Backend Developer", "skills": ["Python", "API"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "IT Infrastructure Upgrade", "alloc": 90, "avail": "10%", "status": "Allocated"},
        {"id": 6, "ini": "MT", "av_col": "#0284C7", "name": "Michael Tan", "email": "michael.tan@hpe.com", "role": "QA Engineer", "skills": ["Testing", "Automation"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "IT Infrastructure Upgrade", "alloc": 120, "avail": "-20%", "status": "Overallocated"},
        {"id": 7, "ini": "DS", "av_col": "#00875A", "name": "Dana Scott", "email": "dana.scott@hpe.com", "role": "Content Specialist", "skills": ["Content", "SEO"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "Marketing Campaign", "alloc": 40, "avail": "60%", "status": "Available"},
        {"id": 8, "ini": "KL", "av_col": "#0284C7", "name": "Karen Lim", "email": "karen.lim@hpe.com", "role": "DevOps Engineer", "skills": ["Azure", "CI/CD"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "IT Infrastructure Upgrade", "alloc": 100, "avail": "0%", "status": "Allocated"},
        {"id": 9, "ini": "NP", "av_col": "#475569", "name": "Noah Park", "email": "noah.park@hpe.com", "role": "IT Support", "skills": ["Infrastructure", "Support"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#F1F5F9", "#334155")], "project": "Internal Operations", "alloc": 60, "avail": "40%", "status": "Available"},
        {"id": 10, "ini": "CT", "av_col": "#EA580C", "name": "Chloe Tan", "email": "chloe.tan@hpe.com", "role": "Data Analyst", "skills": ["Power BI", "Excel"], "sk_cols": [("#DBEAFE", "#1D4ED8"), ("#DCFCE7", "#15803D")], "project": "Sustainability Initiative", "alloc": 100, "avail": "0%", "status": "Allocated"},
    ]

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

        st.multiselect(
            "Resource(s) *",
            ["James Miller", "Sarah Kim", "Alice Morgan", "Ryan Cruz", "Michael Tan", "Emily Rodriguez", "Dana Scott"],
            default=["James Miller", "Sarah Kim"],
            key="ae_res",
        )

        dt1, dt2, dt3, dt4 = st.columns([1.2, 0.9, 1.2, 0.9])
        with dt1:
            ev_start_d = st.date_input("Start Date & Time *", date(2024, 5, 9), key="ae_sd")
        with dt2:
            st.markdown("<div style='font-size:12.5px;margin-bottom:4px'>&nbsp;</div>", unsafe_allow_html=True)
            st.selectbox("Start Time", ["09:00 AM", "10:00 AM", "01:00 PM", "02:00 PM"], index=0, key="ae_st", label_visibility="collapsed")
        with dt3:
            ev_end_d = st.date_input("End Date & Time *", date(2024, 5, 9), key="ae_ed")
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
                log_event("Created", "Calendar Event", None, ev_title or ev_type)
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
        cal_res_f = st.selectbox("Resource", ["All Resources", "James Miller", "Sarah Kim", "Alice Morgan", "Ryan Cruz", "Michael Tan"], key="cal_f_res", label_visibility="collapsed")
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
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;justify-content:center;gap:12px;border:1px solid #DCE8F3;border-radius:8px;padding:6px 10px;background:white">
                    <span style="font-weight:700;color:#475569;cursor:pointer">‹</span>
                    <span style="font-size:14.5px;font-weight:800;color:#0F172A">May 2024</span>
                    <span style="font-weight:700;color:#475569;cursor:pointer">›</span>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
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
        cal_weeks = [
            [
                ("28", False, []),
                ("29", False, []),
                ("30", False, []),
                ("1", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g"), ("● Design Review", "2:00 PM – 4:00 PM", "b")]),
                ("2", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("3", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("4", True, []),
            ],
            [
                ("5", True, []),
                ("6", True, [("● Marketing Campaign", "9:00 AM – 5:00 PM<br>(30%)", "g")]),
                ("7", True, [("● Marketing Campaign", "9:00 AM – 5:00 PM<br>(30%)", "g")]),
                ("8", True, [("● Marketing Campaign", "9:00 AM – 5:00 PM<br>(30%)", "g")]),
                ("9", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g"), ("● Team Meeting", "10:00 AM – 11:00 AM", "b")]),
                ("10", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("11", True, []),
            ],
            [
                ("12", True, []),
                ("13", True, [("● IT Infrastructure", "9:00 AM – 5:00 PM<br>(20%)", "g")]),
                ("14", True, [("● IT Infrastructure", "9:00 AM – 5:00 PM<br>(20%)", "g")]),
                ("15", True, [("● IT Infrastructure", "9:00 AM – 5:00 PM<br>(20%)", "g"), ("● Documentation Review", "2:00 PM – 3:00 PM", "b")]),
                ("16", True, [("● IT Infrastructure", "9:00 AM – 5:00 PM<br>(20%)", "g")]),
                ("17", True, [("● IT Infrastructure", "9:00 AM – 5:00 PM<br>(20%)", "g")]),
                ("18", True, []),
            ],
            [
                ("19", True, []),
                ("20", True, [("📅 Annual Leave", "(Full Day)", "r")]),
                ("21", True, [("📅 Annual Leave", "(Full Day)", "r")]),
                ("22", True, [("📅 Annual Leave", "(Full Day)", "r")]),
                ("23", True, []),
                ("24", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("25", True, []),
            ],
            [
                ("26", True, []),
                ("27", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("28", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("29", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("30", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("31", True, [("● Website Revamp", "9:00 AM – 5:00 PM<br>(50%)", "g")]),
                ("1", False, []),
            ],
        ]

        ev_box_tones = {
            "g": ("#DCFCE7", "#065F46"),
            "b": ("#DBEAFE", "#1D4ED8"),
            "r": ("#FEE2E2", "#B91C1C"),
        }

        cal_rows_html = ""
        for wk in cal_weeks:
            row_cells = ""
            for d_str, in_m, evs in wk:
                is_sel = in_m and d_str == "9"
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
        st.date_input("Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="rep_top_dates", label_visibility="collapsed")
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
            kpi_card_mockup("Total Resources", "24", icon_html="👥", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with rk2:
            kpi_card_mockup("Available", "18", sub_text="75%", icon_html="👤", icon_bg="#D1FAE5", icon_fg="#00875A")
        with rk3:
            kpi_card_mockup("Partial Allocation", "4", sub_text="17%", icon_html="◔", icon_bg="#FEF3C7", icon_fg="#F28C00")
        with rk4:
            kpi_card_mockup("On Leave", "2", sub_text="8%", icon_html="👤", icon_bg="#FEE2E2", icon_fg="#EF4444")

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

            mlist_rows = [
                ("UI/UX Design Complete", "Website Revamp", "Design", "Completed", "Apr 1, 2024", "Apr 30, 2024", "Apr 28, 2024", "—", "Design phase completed."),
                ("Campaign Launch", "Marketing Campaign", "Marketing", "In Progress", "Apr 15, 2024", "May 20, 2024", "—", "Content Review", "Awaiting final assets."),
                ("Infrastructure Setup", "IT Infrastructure Upgrade", "IT", "In Progress", "May 1, 2024", "May 15, 2024", "—", "Requirements Gathering", "Environment setup ongoing."),
                ("Sustainability Plan", "Sustainability Initiative", "Operations", "Not Started", "May 5, 2024", "May 25, 2024", "—", "—", "Pending stakeholder approval."),
                ("Process Documentation", "Internal Operations", "Operations", "Not Started", "May 10, 2024", "May 30, 2024", "—", "Infrastructure Setup", "Template to be finalized."),
                ("QA Testing", "Website Revamp", "IT", "Overdue", "Apr 1, 2024", "Apr 20, 2024", "—", "Development Complete", "Testing issues to resolve."),
                ("Training Materials", "Marketing Campaign", "Marketing", "Overdue", "Apr 20, 2024", "May 5, 2024", "—", "Content Review", "Update with latest product info."),
                ("System Configuration", "IT Infrastructure Upgrade", "IT", "Completed", "Mar 15, 2024", "Apr 10, 2024", "Apr 9, 2024", "—", "Configuration completed successfully."),
            ]
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
                    st.date_input("Filter Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="cr_flt_dates", label_visibility="collapsed")

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
        custom_proj_df = pd.DataFrame(
            [
                {"Project Name": "Website Revamp", "Project Status": "On Track", "Total Tasks": 40, "Completed": 18, "In Progress": 12, "Not Started": 6, "Overdue": 4, "Total Milestones": 8, "Team Members": 8},
                {"Project Name": "Marketing Campaign", "Project Status": "At Risk", "Total Tasks": 36, "Completed": 14, "In Progress": 10, "Not Started": 8, "Overdue": 4, "Total Milestones": 6, "Team Members": 6},
                {"Project Name": "IT Infrastructure Upgrade", "Project Status": "At Risk", "Total Tasks": 28, "Completed": 12, "In Progress": 8, "Not Started": 6, "Overdue": 2, "Total Milestones": 5, "Team Members": 5},
                {"Project Name": "Sustainability Initiative", "Project Status": "Delayed", "Total Tasks": 22, "Completed": 10, "In Progress": 6, "Not Started": 4, "Overdue": 2, "Total Milestones": 4, "Team Members": 4},
                {"Project Name": "Internal Operations", "Project Status": "On Track", "Total Tasks": 18, "Completed": 8, "In Progress": 6, "Not Started": 4, "Overdue": 0, "Total Milestones": 2, "Team Members": 3},
            ]
        )
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

        um_users = [
            ("JM", "#0284C7", "James Miller", "james.miller@hpe.com", "Admin", "IT", "Active"),
            ("SK", "#0284C7", "Sarah Kim", "sarah.kim@hpe.com", "Project Manager", "Design", "Active"),
            ("AL", "#0F172A", "Alice Morgan", "alice.morgan@hpe.com", "Team Member", "Marketing", "Active"),
            ("RJ", "#0F172A", "Ryan Cruz", "ryan.cruz@hpe.com", "Team Member", "Operations", "Active"),
            ("MT", "#0284C7", "Michael Tan", "michael.tan@hpe.com", "Project Manager", "IT", "Active"),
            ("DP", "#1D4ED8", "Daniel Perez", "daniel.perez@hpe.com", "Team Member", "Operations", "Active"),
            ("ER", "#EF4444", "Ella Ramirez", "ella.ramirez@hpe.com", "Team Member", "Design", "Inactive"),
            ("CP", "#0284C7", "Chris Padilla", "chris.padilla@hpe.com", "Team Member", "IT", "Active"),
        ]
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

    st.write("")
    with st.container(border=True):
        st.markdown("<div style='font-size:15.5px;font-weight:800;color:#0F172A'>Data Management</div><div style='font-size:12px;color:#64748B;margin-bottom:10px'>Import the Portfolio worksheet or selectively remove stored database data.</div>", unsafe_allow_html=True)
        dm_left, dm_right = st.columns(2)
        with dm_left:
            st.markdown("**Import Excel Portfolio**")
            portfolio_file = st.file_uploader("Excel workbook", type=["xlsx"], key="settings_portfolio_upload", help="The workbook must include a Portfolio worksheet.")
            import_mode = st.selectbox("Import mode", ["Replace portfolio data", "Append portfolio data"], key="settings_import_mode")
            if portfolio_file is not None:
                try:
                    preview_payload = extract_portfolio_workbook(portfolio_file)
                    st.caption(f"Ready: {len(preview_payload['projects'])} projects, {len(preview_payload['tasks'])} tasks, {len(preview_payload['milestones'])} milestones, {len(preview_payload['resources'])} resources.")
                    if st.button("Import and extract", type="primary", key="settings_import_excel"):
                        counts=import_portfolio(preview_payload, replace_existing=import_mode.startswith("Replace"))
                        st.session_state["settings_import_result"]=counts
                        st.rerun()
                except Exception as exc:
                    st.error(f"Excel import could not be prepared: {exc}")
            if st.session_state.get("settings_import_result"):
                c=st.session_state.pop("settings_import_result")
                st.success(f"Imported {c['projects']} projects, {c['tasks']} tasks, {c['milestones']} milestones, and {c['resources']} resources.")
        with dm_right:
            st.markdown("**Clear saved data**")
            delete_choices = st.multiselect("Choose what to delete", ["Projects","Tasks","Resources","Milestones","Saved reports","Audit log","Settings"], key="settings_delete_choices")
            confirm_clear = st.checkbox("I understand the selected data will be permanently deleted.", key="settings_clear_confirm")
            if st.button("Clear selected data", key="settings_clear_selected", disabled=not(delete_choices and confirm_clear)):
                clear_selected_data(delete_choices)
                st.session_state["settings_clear_result"] = ", ".join(delete_choices)
                st.rerun()
            if st.session_state.get("settings_clear_result"):
                st.success("Cleared: " + st.session_state.pop("settings_clear_result"))
    if st.session_state.get("settings_modal") == "add_user":
        add_user_settings_modal()


# ---------------------------------------------------------------------------
# Deployment notes and lightweight runtime integrity checks
# ---------------------------------------------------------------------------
# Configure Streamlit Secrets (or environment variable) without committing credentials:
# .streamlit/secrets.toml
# MONGODB_URI = "mongodb+srv://<db-user>:<password>@<cluster-host>/?retryWrites=true&w=majority"
# The URI is intentionally never embedded in this source file.
