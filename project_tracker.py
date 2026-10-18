"""Vektor Project Management - High-Fidelity Streamlit Application.
Install: pip install streamlit pandas plotly openpyxl
Run: streamlit run vektor_project_management_app.py
Uses a local SQLite database (vektor.db), seeded with demo data on first run.
"""
from __future__ import annotations

import base64
import calendar
import html
import io
import json
import os
import sqlite3
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DB = os.environ.get("VEKTOR_DB_PATH", "vektor.db")

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
    c = sqlite3.connect(DB, check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c


def sql(q: str, p: tuple = ()) -> None:
    with conn() as c:
        c.execute(q, p)
        c.commit()


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

        if c.execute("SELECT COUNT(*) FROM projects").fetchone()[0] == 0:
            # Exact 12 projects from the uploaded mockups (List View order #1 to #12)
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

html, body, [class*="css"] {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: {TEXT_COLOR};
}}
.stApp {{
    background: {APP_BG};
    color: {TEXT_COLOR};
}}
.block-container {{
    padding-top: 1.2rem !important;
    padding-bottom: 2rem !important;
    padding-left: 1.8rem !important;
    padding-right: 1.8rem !important;
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
    font-size: 12.8px;
    text-align: left;
}}
table.vk-table thead tr {{
    background: #F8FAFC;
    border-bottom: 1px solid #E2E8F0;
}}
table.vk-table th {{
    padding: 9px 12px;
    font-weight: 600;
    font-size: 11.5px;
    color: #475569;
}}
table.vk-table td {{
    padding: 10px 12px;
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
            circle = '<div style="width:24px;height:24px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;z-index:2">✓</div>'
        elif state == "current":
            circle = '<div style="width:24px;height:24px;border-radius:50%;background:#1473E6;border:3px solid #BFDBFE;color:white;display:flex;align-items:center;justify-content:center;font-size:10px;z-index:2">●</div>'
        else:
            circle = '<div style="width:22px;height:22px;border-radius:50%;background:#94A3B8;color:white;display:flex;align-items:center;justify-content:center;font-size:10px;z-index:2"></div>'

        line_color = "#00875A" if idx < 2 else ("#1473E6" if idx == 2 else "#CBD5E1")
        line_right = (
            f'<div style="position:absolute;top:46px;left:50%;width:100%;height:3px;background:{line_color};z-index:1"></div>'
            if idx < len(steps) - 1
            else ""
        )
        nodes_html += f"""
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

    # Dark Teal Header Banner
    st.markdown(
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
        """,
        unsafe_allow_html=True,
    )

    # 5 Info Cards Row
    st.markdown(
        f"""
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
        """,
        unsafe_allow_html=True,
    )

    # Row: Project Progress + Timeline
    c_left, c_right = st.columns([1.05, 1.95])
    with c_left, st.container(border=True):
        section_title("📄 Project Progress")
        d1, d2 = st.columns([1, 1.15])
        with d1:
            st.plotly_chart(render_project_progress_donut(pct, 165), use_container_width=True, config={"displayModeBar": False})
        with d2:
            st.markdown(
                """
                <div style="font-size:12px;line-height:2.1;padding-top:14px">
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Completed</span><b>14 (70%)</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> In Progress</span><b>4 (20%)</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> On Hold</span><b>1 (5%)</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#64748B">●</span> Not Started</span><b>1 (5%)</b></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    with c_right, st.container(border=True):
        section_title("⏳ Timeline", "View Full Timeline →")
        st.markdown(render_timeline_stepper_html(), unsafe_allow_html=True)

    # Bottom Row: Project Information + Key Milestones & Recent Activity
    b_left, b_right = st.columns([1.05, 1.95])
    with b_left, st.container(border=True):
        section_title("📄 Project Information")
        st.markdown(
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
            """,
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
        f"""
        <div style="background:{DARK_TEAL_HEADER};color:white;padding:16px 22px;border-radius:10px;margin-bottom:14px;display:flex;align-items:center;gap:12px">
            <div style="font-size:26px">📝</div>
            <div>
                <div style="font-size:19px;font-weight:800;color:white">{title_txt}</div>
                <div style="font-size:12px;color:#CBD5E1">{sub_txt}</div>
            </div>
        </div>
        """,
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
            """
            <div style="font-size:12px;line-height:2.1;padding:4px 8px">
                <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Completed</span><b>14 (70%)</b></div>
                <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> In Progress</span><b>4 (20%)</b></div>
                <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> On Hold</span><b>1 (5%)</b></div>
                <div style="display:flex;justify-content:space-between"><span><span style="color:#64748B">●</span> Not Started</span><b>1 (5%)</b></div>
            </div>
            """,
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
        """,
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
# PAGE 3: GANTT CHART
# ===========================================================================
elif page == "Gantt Chart":
    render_top_header("Gantt Chart", "Visualize project schedules, task durations, dependencies, and milestones", show_filters=True)
    t = apply_global_filters(T, "project")
    m = apply_global_filters(M, "project")

    with st.container(border=True):
        c1, c2, c3 = st.columns([1.4, 1, 1])
        with c1:
            view_mode = tile_toggle("Timeline View", ["Tasks & Milestones", "Tasks Only", "Projects Timeline"], key="gantt_view_mode")
        with c2:
            color_by = st.selectbox("Color Bars By", ["status", "priority", "project", "team"])
        with c3:
            st.write("")
            show_deps = st.checkbox("Show Dependency Labels on Bars", value=True)

        if view_mode == "Projects Timeline":
            g = apply_global_filters(P, "name").copy()
            g["Start"] = pd.to_datetime(g.start_date, errors="coerce")
            g["Finish"] = pd.to_datetime(g.end_date, errors="coerce")
            y_col = "name"
        else:
            g = t.copy()
            g["Start"] = pd.to_datetime(g.start_date, errors="coerce")
            g["Finish"] = pd.to_datetime(g.due_date, errors="coerce")
            y_col = "name"

        g = g.dropna(subset=["Start", "Finish"])
        if g.empty:
            st.info("No scheduled items available for the selected filters.")
        else:
            cmap = {
                "Completed": GREEN,
                "On Track": GREEN,
                "In Progress": BLUE,
                "At Risk": ORANGE,
                "Not Started": GRAY,
                "Overdue": RED,
                "Delayed": RED,
                "High": RED,
                "Medium": ORANGE,
                "Low": GREEN,
            }
            g["BarLabel"] = g.apply(
                lambda r: f"{r.get('progress', 0)}%" + (f" ⟵ {r.get('dependency')}" if show_deps and r.get("dependency") else ""),
                axis=1,
            )
            fig = px.timeline(
                g,
                x_start="Start",
                x_end="Finish",
                y=y_col,
                color=color_by if color_by in g.columns else "status",
                text="BarLabel",
                color_discrete_map=cmap,
            )
            if view_mode == "Tasks & Milestones" and not m.empty:
                m_valid = m.dropna(subset=["due_date"]).copy()
                m_valid["Due"] = pd.to_datetime(m_valid.due_date, errors="coerce")
                fig.add_trace(
                    go.Scatter(
                        x=m_valid["Due"],
                        y=m_valid["name"],
                        mode="markers",
                        marker=dict(symbol="diamond", size=12, color=NAVY, line=dict(width=2, color="white")),
                        name="Milestone Deadline",
                        hovertext=m_valid["project"] + " • " + m_valid["status"],
                    )
                )
            fig.update_yaxes(autorange="reversed", title=None, showgrid=True, gridcolor="#F1F5F9")
            fig.update_xaxes(showgrid=True, gridcolor="#E2E8F0")
            fig.update_layout(
                height=max(380, 42 * len(g)),
                margin=dict(l=0, r=16, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                legend_title_text="",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with st.container(border=True):
        section_title("Schedule & Dependency Matrix", f"{len(t)} Scheduled Tasks")
        render_rich_table(
            t,
            [
                ("name", "Task", "bold"),
                ("project", "Project", "text"),
                ("assignee", "Assignee", "avatar"),
                ("start_date", "Start Date", "date"),
                ("due_date", "Due Date", "date"),
                ("dependency", "Dependency", "text"),
                ("status", "Status", "badge"),
                ("progress", "Progress", "progress"),
            ],
        )


# ===========================================================================
# PAGE 4: TASK LIST
# ===========================================================================
elif page == "Task List":
    render_top_header("Task List", "Manage task scope, assignments, priorities, and delivery dates", show_filters=True)
    t = apply_global_filters(T, "project")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("Total Tasks", len(t), "Filtered scope", "☑️", "blue")
    with k2:
        metric_card("Completed", int((t.status == "Completed").sum()), "Delivered tasks", "✅", "green")
    with k3:
        metric_card("In Progress", int((t.status == "In Progress").sum()), "Active execution", "▶️", "blue")
    with k4:
        metric_card("Not Started / Overdue", int(t.status.isin(["Not Started", "Overdue"]).sum()), "Pending action", "⏸️", "orange")

    st.write("")
    ta, tb = st.tabs(["☷ Task Directory", "＋ Create / Edit Task"])
    with ta:
        with st.container(border=True):
            c1, c2, c3 = st.columns([1.8, 1, 1])
            with c1:
                q = st.text_input("🔍 Search tasks by name, project, or assignee", key="task_search")
            with c2:
                sf = st.selectbox("Status", ["All Statuses", "Completed", "In Progress", "Not Started", "Overdue"], key="task_sf")
            with c3:
                pf_prio = st.selectbox("Priority", ["All Priorities", "High", "Medium", "Low"], key="task_pf")

            v = t.copy()
            if q:
                v = v[v.astype(str).apply(lambda s: s.str.contains(q, case=False, na=False)).any(axis=1)]
            if sf != "All Statuses":
                v = v[v.status == sf]
            if pf_prio != "All Priorities":
                v = v[v.priority == pf_prio]

            section_title("Task List", f"{len(v)} Tasks")
            render_rich_table(
                v,
                [
                    ("name", "Task Name", "bold"),
                    ("project", "Project", "text"),
                    ("team", "Team", "text"),
                    ("status", "Status", "badge"),
                    ("priority", "Priority", "badge"),
                    ("start_date", "Start Date", "date"),
                    ("due_date", "Due Date", "date"),
                    ("assignee", "Assignee", "avatar"),
                    ("progress", "Progress", "progress"),
                    ("dependency", "Dependency", "text"),
                ],
            )
            st.write("")
            st.download_button(
                "⬇ Export Tasks CSV",
                v.to_csv(index=False).encode("utf-8-sig"),
                file_name="vektor_tasks.csv",
                mime="text/csv",
            )

    with tb:
        with st.container(border=True):
            mode = tile_toggle("Action", ["Create Task", "Edit Existing Task"], key="task_action_mode")
            row = None
            if mode == "Edit Existing Task" and not T.empty:
                opts = {f"{r['name']} ({r['project']})": int(r.id) for _, r in T.iterrows()}
                tid = opts[st.selectbox("Select Task", list(opts.keys()))]
                row = T[T.id == tid].iloc[0]

            with st.form("task_crud_form"):
                name = st.text_input("Task Name *", str(row["name"]) if row is not None else "")
                a, b, c = st.columns(3)
                with a:
                    proj_default = row.project if row is not None else (projects_list[0] if projects_list else "")
                    proj = st.selectbox("Project", projects_list, index=projects_list.index(proj_default) if proj_default in projects_list else 0)
                with b:
                    td = str(row.team) if row is not None else teams_list[0]
                    team = st.selectbox("Team", teams_list, index=teams_list.index(td) if td in teams_list else 0)
                with c:
                    statuses = ["Not Started", "In Progress", "Completed", "Overdue"]
                    sd = str(row.status) if row is not None else get_setting("default_task_status", "Not Started")
                    status = st.selectbox("Status", statuses, index=statuses.index(sd) if sd in statuses else 0)

                a, b, c = st.columns(3)
                with a:
                    priorities = ["High", "Medium", "Low"]
                    pr = str(row.priority) if row is not None else get_setting("default_task_priority", "Medium")
                    priority = st.selectbox("Priority", priorities, index=priorities.index(pr) if pr in priorities else 1)
                with b:
                    start = st.date_input("Start Date", date.fromisoformat(row.start_date) if row is not None and row.start_date else date.today())
                with c:
                    due = st.date_input("Due Date", date.fromisoformat(row.due_date) if row is not None and row.due_date else date.today() + timedelta(days=14))

                a, b = st.columns(2)
                with a:
                    assignees = ["Unassigned"] + R.name.tolist()
                    ad = str(row.assignee) if row is not None and row.assignee else "Unassigned"
                    assignee = st.selectbox("Assignee", assignees, index=assignees.index(ad) if ad in assignees else 0)
                with b:
                    dep_opts = ["None"] + [x for x in T.name.tolist() if row is None or x != row["name"]]
                    cur_dep = str(row.dependency) if row is not None and row.dependency else "None"
                    dependency = st.selectbox("Dependency", dep_opts, index=dep_opts.index(cur_dep) if cur_dep in dep_opts else 0)

                progress = st.slider("Progress (%)", 0, 100, int(row.progress or 0) if row is not None else 0)
                notes = st.text_area("Notes", str(row.notes or "") if row is not None else "")
                save = st.form_submit_button("💾 Save Task", type="primary")
                if save:
                    if not name.strip() or due < start:
                        st.error("Task name and valid date range are required.")
                    else:
                        vals = (
                            name.strip(),
                            pmap.get(proj),
                            team,
                            status,
                            priority,
                            start.isoformat(),
                            due.isoformat(),
                            "" if assignee == "Unassigned" else assignee,
                            progress,
                            "" if dependency == "None" else dependency,
                            notes,
                        )
                        if mode == "Create Task":
                            sql(
                                "INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                vals,
                            )
                            log_event("Created", "Task", None, name.strip())
                        else:
                            sql(
                                "UPDATE tasks SET name=?,project_id=?,team=?,status=?,priority=?,start_date=?,due_date=?,assignee=?,progress=?,dependency=?,notes=? WHERE id=?",
                                vals + (int(row.id),),
                            )
                            log_event("Updated", "Task", int(row.id), name.strip())
                        st.success("Task saved.")
                        st.rerun()

            if mode == "Edit Existing Task" and row is not None:
                if st.button("🗑 Delete Selected Task"):
                    sql("DELETE FROM tasks WHERE id=?", (int(row.id),))
                    log_event("Deleted", "Task", int(row.id), row["name"])
                    st.rerun()


# ===========================================================================
# PAGE 5: RESOURCES
# ===========================================================================
elif page == "Resources":
    render_top_header("Resources", "Manage team members, capacity, availability, and project assignments", show_filters=True)
    r = apply_global_filters(R, "none")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("Total Resources", len(r), "Team directory", "👥", "blue")
    with k2:
        metric_card("Available", int((r.status == "Available").sum()), "100% capacity", "🟢", "green")
    with k3:
        metric_card("Partial Allocation", int((r.status == "Partial").sum()), "50% capacity", "🟠", "orange")
    with k4:
        metric_card("On Leave", int((r.status == "On Leave").sum()), "Currently unavailable", "🔴", "red")

    st.write("")
    c1, c2 = st.columns(2)
    with c1, st.container(border=True):
        section_title("Resource Availability", f"{len(r)} Members")
        chart_status_donut(r, "status", "Members")
    with c2, st.container(border=True):
        section_title("Assignments by Project", "Active Task Distribution")
        cnt = T.groupby("project").size().rename("Assignments").reset_index()
        if not cnt.empty:
            fig = px.bar(
                cnt,
                x="Assignments",
                y="project",
                orientation="h",
                text="Assignments",
                color="Assignments",
                color_continuous_scale=[BLUE, GREEN],
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(
                height=240,
                margin=dict(l=0, r=28, t=6, b=6),
                xaxis_title=None,
                yaxis_title=None,
                coloraxis_showscale=False,
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    ta, tb = st.tabs(["👥 Resource Directory", "＋ Add / Edit Resource"])
    with ta:
        with st.container(border=True):
            q = st.text_input("🔍 Search resources by name, email, role, or team")
            v = r.copy()
            if q:
                v = v[v.astype(str).apply(lambda s: s.str.contains(q, case=False, na=False)).any(axis=1)]
            v["active_tasks"] = v.name.map(T.groupby("assignee").size().to_dict()).fillna(0).astype(int)
            render_rich_table(
                v,
                [
                    ("name", "Team Member", "avatar"),
                    ("email", "Email", "text"),
                    ("team", "Team", "text"),
                    ("role", "Role", "badge"),
                    ("status", "Availability", "badge"),
                    ("capacity", "Capacity", "progress"),
                    ("active_tasks", "Active Tasks", "bold"),
                ],
            )
            st.write("")
            st.download_button(
                "⬇ Export Resources CSV",
                v.to_csv(index=False).encode("utf-8-sig"),
                file_name="vektor_resources.csv",
                mime="text/csv",
            )

    with tb:
        with st.container(border=True):
            mode = tile_toggle("Resource Action", ["Add Resource", "Edit Resource"], key="res_action_mode")
            row = None
            if mode == "Edit Resource" and not R.empty:
                opts = {f"{x['name']} ({x['team']})": int(x.id) for _, x in R.iterrows()}
                rid = opts[st.selectbox("Select Resource", list(opts.keys()))]
                row = R[R.id == rid].iloc[0]
            with st.form("resource_crud_form"):
                a, b = st.columns(2)
                with a:
                    name = st.text_input("Full Name *", str(row["name"]) if row is not None else "")
                with b:
                    email = st.text_input("Email", str(row.email) if row is not None else "")
                a, b, c = st.columns(3)
                with a:
                    td = str(row.team) if row is not None else teams_list[0]
                    team = st.selectbox("Team", teams_list, index=teams_list.index(td) if td in teams_list else 0)
                with b:
                    roles = ["Admin", "Project Manager", "Team Member"]
                    rd = str(row.role) if row is not None else "Team Member"
                    role = st.selectbox("Role", roles, index=roles.index(rd) if rd in roles else 2)
                with c:
                    states = ["Available", "Partial", "On Leave"]
                    sd = str(row.status) if row is not None else "Available"
                    status = st.selectbox("Availability", states, index=states.index(sd) if sd in states else 0)
                capacity = st.slider("Capacity (%)", 0, 100, int(row.capacity or 100) if row is not None else 100)
                notes = st.text_area("Notes", str(row.notes or "") if row is not None else "")
                save = st.form_submit_button("💾 Save Resource", type="primary")
                if save:
                    if not name.strip():
                        st.error("Name is required.")
                    elif mode == "Add Resource":
                        sql(
                            "INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",
                            (name.strip(), email.strip(), team, role, status, capacity, notes),
                        )
                        log_event("Created", "Resource", None, name.strip())
                        st.rerun()
                    else:
                        sql(
                            "UPDATE resources SET name=?,email=?,team=?,role=?,status=?,capacity=?,notes=? WHERE id=?",
                            (name.strip(), email.strip(), team, role, status, capacity, notes, int(row.id)),
                        )
                        log_event("Updated", "Resource", int(row.id), name.strip())
                        st.rerun()


# ===========================================================================
# PAGE 6: CALENDAR
# ===========================================================================
elif page == "Calendar":
    render_top_header("Calendar", "View project dates, task due dates, and milestone deadlines", show_filters=True)
    t = apply_global_filters(T, "project")
    m = apply_global_filters(M, "project")

    with st.container(border=True):
        c1, _ = st.columns([1, 2.5])
        with c1:
            month_pick = st.date_input("Select Calendar Month", date(2024, 5, 1))
        first = month_pick.replace(day=1)
        last = month_pick.replace(day=calendar.monthrange(month_pick.year, month_pick.month)[1])

        events = []
        for _, r in t.iterrows():
            try:
                d = date.fromisoformat(str(r.due_date))
            except Exception:
                continue
            if first <= d <= last:
                events.append({"Date": d, "Type": "Task", "Title": r["name"], "Project": r["project"], "Status": r.status})
        for _, r in m.iterrows():
            try:
                d = date.fromisoformat(str(r.due_date))
            except Exception:
                continue
            if first <= d <= last:
                events.append({"Date": d, "Type": "Milestone", "Title": r["name"], "Project": r["project"], "Status": r.status})

        ev = pd.DataFrame(events)
        section_title(first.strftime("%B %Y"), f"{len(ev)} Scheduled Deadlines")

        weeks = calendar.Calendar(firstweekday=6).monthdatescalendar(first.year, first.month)
        hdr_cols = st.columns(7)
        for col, dayname in zip(hdr_cols, ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]):
            col.markdown(f"<div style='text-align:center;font-weight:700;font-size:12px;color:#475569;padding:4px'>{dayname}</div>", unsafe_allow_html=True)

        for week in weeks:
            cols = st.columns(7)
            for col, d in zip(cols, week):
                with col:
                    in_month = d.month == first.month
                    day_color = TEXT_COLOR if in_month else "#CBD5E1"
                    daily = ev[ev.Date == d] if not ev.empty else pd.DataFrame()
                    pills = ""
                    for _, e in daily.head(3).iterrows():
                        bg = GREEN if e["Type"] == "Milestone" else BLUE
                        icon = "⚑" if e["Type"] == "Milestone" else "☑"
                        pills += f'<div style="background:{bg};color:white;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{icon} {html.escape(e["Title"])}</div>'
                    st.markdown(
                        f'<div style="min-height:86px;border:1px solid {BORDER_COLOR};border-radius:8px;padding:6px;background:{"#FFFFFF" if in_month else "#F8FAFC"};margin-bottom:6px">'
                        f'<div style="font-size:12px;font-weight:700;color:{day_color}">{d.day}</div>'
                        f"{pills}</div>",
                        unsafe_allow_html=True,
                    )

    with st.container(border=True):
        section_title("Deadlines in Selected Month", first.strftime("%B %Y"))
        if ev.empty:
            st.info("No tasks or milestones due in this month. Try selecting April, May, or June 2024.")
        else:
            render_rich_table(
                ev.sort_values("Date"),
                [
                    ("Date", "Due Date", "date"),
                    ("Type", "Item Type", "bold"),
                    ("Title", "Name", "bold"),
                    ("Project", "Project", "text"),
                    ("Status", "Status", "badge"),
                ],
            )


# ===========================================================================
# PAGE 7: REPORTS
# ===========================================================================
elif page == "Reports":
    render_top_header("Reports", "Gain insights into project progress, resource utilization, and team performance.", show_filters=False)

    with st.container(border=True):
        f1, f2, f3, f4, f5 = st.columns([1.45, 1.05, 1.05, 0.95, 0.9])
        with f1:
            dates = st.date_input("Date Range", (date(2024, 4, 24), date(2024, 6, 30)), key="report_dates")
        with f2:
            pf = st.selectbox("Project", ["All Projects"] + projects_list, key="report_project")
        with f3:
            tf = st.selectbox("Team", ["All Teams"] + teams_list, key="report_team")
        with f4:
            fmt = st.selectbox("Export Format", ["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"], key="report_fmt")

        rp = P.copy()
        rt = T.copy()
        rr = R.copy()
        rm = M.copy()
        if pf != "All Projects":
            rp = rp[rp.name == pf]
            rt = rt[rt.project == pf]
            rm = rm[rm.project == pf]
        if tf != "All Teams":
            rp = rp[rp.team == tf]
            rt = rt[rt.team == tf]
            rr = rr[rr.team == tf]
            rm = rm[rm.team == tf]
        if isinstance(dates, tuple) and len(dates) == 2:
            ds, de = dates[0].isoformat(), dates[1].isoformat()
            rt = rt[(rt.start_date.fillna("") <= de) & (rt.due_date.fillna("") >= ds)]
            rm = rm[(rm.due_date.fillna("") >= "2024-01-01") & (rm.due_date.fillna("") <= de)]

        with f5:
            st.write("")
            if fmt.startswith("Excel"):
                st.download_button(
                    "⬇ Export Report",
                    excel_bytes({"Projects": rp, "Tasks": rt, "Resources": rr, "Milestones": rm}),
                    file_name="vektor_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            elif fmt.startswith("CSV"):
                st.download_button(
                    "⬇ Export Report",
                    rp.to_csv(index=False).encode("utf-8-sig"),
                    file_name="vektor_report.csv",
                    mime="text/csv",
                    type="primary",
                    use_container_width=True,
                )
            else:
                payload = {
                    k: v.where(pd.notna(v), None).to_dict("records")
                    for k, v in {"Projects": rp, "Tasks": rt, "Resources": rr, "Milestones": rm}.items()
                }
                st.download_button(
                    "⬇ Export Report",
                    json.dumps(payload, default=str, indent=2).encode(),
                    file_name="vektor_report.json",
                    mime="application/json",
                    type="primary",
                    use_container_width=True,
                )

    tabs = st.tabs(["Overview", "Projects", "Resources", "Tasks", "Milestones", "Workload", "Custom"])

    with tabs[0]:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Total Projects", len(rp), f"{int((rp.status == 'On Track').sum())} on track", "📁", "green")
        with k2:
            metric_card("Total Tasks", len(rt), f"{int((rt.status == 'Completed').sum())} completed", "☑️", "blue")
        with k3:
            metric_card("Milestones", len(rm), f"{int((rm.status == 'Completed').sum())} achieved", "⚑", "orange")
        with k4:
            metric_card("Team Members", len(rr), f"{int((rr.status == 'Available').sum())} available", "👥", "purple")

        st.write("")
        c1, c2, c3 = st.columns(3)
        with c1, st.container(border=True):
            section_title("Project Status")
            chart_status_donut(rp, "status", "Projects")
        with c2, st.container(border=True):
            section_title("Task Status")
            chart_status_donut(rt, "status", "Tasks")
        with c3, st.container(border=True):
            section_title("Resource Availability")
            chart_status_donut(rr, "status", "Members")

        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            section_title("Project Progress")
            progress_bar_chart(rp.sort_values("progress"), "name", "progress")
        with c2, st.container(border=True):
            section_title("Task Completion Trend")
            trend = pd.DataFrame(
                {
                    "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                    "Completed": [8, 18, 27, 35, 40, 48],
                    "In Progress": [3, 12, 16, 19, 21, 24],
                }
            )
            fig = px.line(
                trend,
                x="Month",
                y=["Completed", "In Progress"],
                markers=True,
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE},
            )
            fig.update_layout(height=250, margin=dict(l=0, r=0, t=6, b=6), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with tabs[1]:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Total Projects", len(rp), "All portfolio projects", "📁", "green")
        with k2:
            metric_card("On Track", int((rp.status == "On Track").sum()), "Healthy delivery", "✅", "green")
        with k3:
            metric_card("In Progress", int(((rp.progress > 0) & (rp.progress < 100)).sum()), "Active execution", "▶️", "blue")
        with k4:
            metric_card("At Risk / Delayed", int(rp.status.isin(["At Risk", "Delayed"]).sum()), "Needs mitigation", "⚠️", "red")

        st.write("")
        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            section_title("Project Status")
            chart_status_donut(rp, "status", "Projects")
        with c2, st.container(border=True):
            section_title("Project Progress")
            progress_bar_chart(rp.sort_values("progress"), "name", "progress")

        with st.container(border=True):
            section_title("Project Report Summary", f"{len(rp)} Projects")
            render_rich_table(
                rp,
                [
                    ("name", "Project Name", "bold"),
                    ("owner", "Owner", "avatar"),
                    ("team", "Team", "text"),
                    ("status", "Status", "badge"),
                    ("start_date", "Start Date", "date"),
                    ("end_date", "End Date", "date"),
                    ("progress", "Progress", "progress"),
                    ("health", "Health", "badge"),
                ],
            )

    with tabs[2]:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Total Resources", len(rr), "Workspace members", "👥", "blue")
        with k2:
            metric_card("Available", int((rr.status == "Available").sum()), "Ready for assignment", "🟢", "green")
        with k3:
            metric_card("Partial", int((rr.status == "Partial").sum()), "Part-time allocation", "🟠", "orange")
        with k4:
            metric_card("On Leave", int((rr.status == "On Leave").sum()), "Out of office", "🔴", "red")

        st.write("")
        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            section_title("Resource Status")
            chart_status_donut(rr, "status", "Members")
        with c2, st.container(border=True):
            section_title("Team Distribution")
            td = rr.team.value_counts().rename_axis("Team").reset_index(name="Members")
            fig = px.bar(td, x="Team", y="Members", text="Members", color="Team", color_discrete_sequence=[GREEN, BLUE, ORANGE, PURPLE, GRAY])
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(height=240, showlegend=False, margin=dict(l=0, r=0, t=12, b=0), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        with st.container(border=True):
            section_title("Resource Directory & Capacity", f"{len(rr)} Members")
            rr_view = rr.copy()
            rr_view["tasks"] = rr_view.name.map(T.groupby("assignee").size().to_dict()).fillna(0).astype(int)
            render_rich_table(
                rr_view,
                [
                    ("name", "Resource Name", "avatar"),
                    ("role", "Role", "badge"),
                    ("team", "Team", "text"),
                    ("status", "Status", "badge"),
                    ("capacity", "Capacity", "progress"),
                    ("tasks", "Assigned Tasks", "bold"),
                ],
            )

    with tabs[3]:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Total Tasks", len(rt), "In date range", "☑️", "blue")
        with k2:
            metric_card("Completed", int((rt.status == "Completed").sum()), "100% done", "✅", "green")
        with k3:
            metric_card("In Progress", int((rt.status == "In Progress").sum()), "Active work", "▶️", "blue")
        with k4:
            metric_card("Not Started", int((rt.status == "Not Started").sum()), "Upcoming", "⏸️", "orange")

        st.write("")
        c1, c2, c3 = st.columns(3)
        with c1, st.container(border=True):
            section_title("Task Status")
            chart_status_donut(rt, "status", "Tasks")
        with c2, st.container(border=True):
            section_title("Tasks by Project")
            z = rt.groupby("project").size().rename("Tasks").reset_index()
            fig = px.bar(z, x="Tasks", y="project", orientation="h", text="Tasks", color_discrete_sequence=[BLUE])
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(height=240, margin=dict(l=0, r=24, t=6, b=6), xaxis_title=None, yaxis_title=None, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with c3, st.container(border=True):
            section_title("Tasks by Priority")
            chart_status_donut(rt, "priority", "Priority")

        with st.container(border=True):
            section_title("Tasks Report Table", f"{len(rt)} Tasks")
            render_rich_table(
                rt,
                [
                    ("name", "Task Name", "bold"),
                    ("project", "Project", "text"),
                    ("team", "Team", "text"),
                    ("status", "Status", "badge"),
                    ("priority", "Priority", "badge"),
                    ("start_date", "Start Date", "date"),
                    ("due_date", "Due Date", "date"),
                    ("assignee", "Assignee", "avatar"),
                    ("progress", "Progress", "progress"),
                    ("dependency", "Dependency", "text"),
                ],
            )

    with tabs[4]:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Total Milestones", len(rm), "Tracked checkpoints", "⚑", "blue")
        with k2:
            metric_card("Completed", int((rm.status == "Completed").sum()), "Delivered on time", "✅", "green")
        with k3:
            metric_card("In Progress", int((rm.status == "In Progress").sum()), "Active milestones", "▶️", "blue")
        with k4:
            metric_card("Overdue", int((rm.status == "Overdue").sum()), "Past due date", "⚠️", "red")

        st.write("")
        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            section_title("Milestone Status")
            chart_status_donut(rm, "status", "Milestones")
        with c2, st.container(border=True):
            section_title("Milestones by Project")
            z = rm.groupby("project").size().rename("Milestones").reset_index()
            fig = px.bar(z, x="project", y="Milestones", text="Milestones", color="project", color_discrete_sequence=[GREEN, BLUE, ORANGE, RED, PURPLE])
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(height=240, showlegend=False, margin=dict(l=0, r=0, t=14, b=6), xaxis_title=None, yaxis_title=None, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        with st.container(border=True):
            section_title("Milestones Report Table", f"{len(rm)} Milestones")
            render_rich_table(
                rm,
                [
                    ("name", "Milestone", "bold"),
                    ("project", "Project", "text"),
                    ("team", "Team", "text"),
                    ("status", "Status", "badge"),
                    ("start_date", "Start Date", "date"),
                    ("due_date", "Due Date", "date"),
                    ("completion_date", "Completion Date", "date"),
                    ("dependency", "Dependency", "text"),
                    ("notes", "Notes", "text"),
                ],
            )

    with tabs[5]:
        counts = T.groupby("assignee").size().rename("active_tasks").reset_index()
        work = rr.merge(counts, left_on="name", right_on="assignee", how="left").drop(columns=["assignee"], errors="ignore")
        work["active_tasks"] = work["active_tasks"].fillna(0).astype(int)
        work["utilization"] = work.apply(
            lambda r: 0 if r["status"] == "On Leave" else min(100, int(r["active_tasks"] * 35 + (55 if r["status"] == "Available" else 30))),
            axis=1,
        )
        work["workload_status"] = work["utilization"].apply(
            lambda u: "Overloaded" if u >= 95 else ("High" if u >= 75 else "Balanced")
        )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card("Team Members", len(work), "Evaluated resources", "👥", "blue")
        with k2:
            metric_card("Active Tasks", int(work["active_tasks"].sum()), "Assigned across teams", "☑️", "green")
        with k3:
            metric_card("High Workload", int((work["workload_status"] == "High").sum()), "75% – 94% utilization", "🟠", "orange")
        with k4:
            metric_card("Overloaded", int((work["workload_status"] == "Overloaded").sum()), "95%+ utilization", "🔴", "red")

        st.write("")
        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            section_title("Individual Workload", "Top 10 by Utilization %")
            z = work.sort_values("utilization", ascending=True).tail(10)
            fig = px.bar(
                z,
                x="utilization",
                y="name",
                orientation="h",
                text="utilization",
                color="workload_status",
                color_discrete_map={"Balanced": GREEN, "High": ORANGE, "Overloaded": RED},
            )
            fig.update_traces(texttemplate="%{text}%", textposition="outside", cliponaxis=False)
            fig.update_layout(height=300, margin=dict(l=0, r=36, t=6, b=6), xaxis_title=None, yaxis_title=None, legend_title_text="", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with c2, st.container(border=True):
            section_title("Workload by Team", "Total Active Tasks")
            z = work.groupby("team")["active_tasks"].sum().reset_index()
            fig = px.bar(z, x="team", y="active_tasks", text="active_tasks", color="team", color_discrete_sequence=[GREEN, BLUE, ORANGE, PURPLE, GRAY])
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(height=300, showlegend=False, margin=dict(l=0, r=0, t=14, b=6), xaxis_title=None, yaxis_title=None, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        with st.container(border=True):
            section_title("Resource Workload Breakdown", f"{len(work)} Members")
            render_rich_table(
                work.sort_values("utilization", ascending=False),
                [
                    ("name", "Team Member", "avatar"),
                    ("role", "Role", "badge"),
                    ("team", "Team", "text"),
                    ("status", "Availability", "badge"),
                    ("active_tasks", "Active Tasks", "bold"),
                    ("utilization", "Utilization", "progress"),
                    ("workload_status", "Workload Status", "badge"),
                ],
            )

    with tabs[6]:
        with st.container(border=True):
            section_title("Build Custom Report", "Configure dimensions, metrics, and visualization style")
            a, b, c = st.columns(3)
            with a:
                typ = st.selectbox("Report Type", ["Project Report", "Task Report", "Resource Report", "Milestone Report", "Workload Report"])
            with b:
                group = st.selectbox("Group By", ["Status", "Team", "Project", "Priority", "Assignee"])
            with c:
                metrics = st.multiselect(
                    "Metrics / Data to Include",
                    ["Project Status", "Task Status", "Milestone Status", "Team Members", "Progress", "Priority", "Dates", "Assignees"],
                    default=["Project Status", "Task Status", "Progress"],
                )
            viz = tile_toggle(
                "Visualization Type",
                ["Bar Chart", "Column Chart", "Pie Chart", "Line Chart", "Table", "Summary Cards"],
                key="custom_rep_viz",
            )
            b1, b2, _ = st.columns([0.8, 0.8, 3])
            with b1:
                generate = st.button("▶ Generate Report", type="primary", use_container_width=True)
            with b2:
                reset = st.button("↺ Reset", use_container_width=True)

            if reset:
                st.session_state.pop("custom_report", None)
                st.rerun()
            if generate or "custom_report" not in st.session_state:
                st.session_state.custom_report = {"type": typ, "group": group, "metrics": metrics, "viz": viz}

        if st.session_state.get("custom_report"):
            cfg = st.session_state.custom_report
            source_map = {
                "Project Report": rp,
                "Task Report": rt,
                "Resource Report": rr,
                "Milestone Report": rm,
                "Workload Report": rr,
            }
            source = source_map.get(cfg["type"], rp).copy()
            colmap = {"Project": "project", "Team": "team", "Status": "status", "Priority": "priority", "Assignee": "assignee"}
            gc = colmap.get(cfg["group"], "status")
            if gc not in source.columns:
                gc = "status" if "status" in source.columns else source.columns[0]
            z = source[gc].fillna("Unassigned").value_counts().rename_axis(gc).reset_index(name="Count")

            with st.container(border=True):
                section_title(f"Report Preview: {cfg['type']} grouped by {cfg['group']}", f"{len(source)} Records")
                if cfg["viz"] == "Pie Chart":
                    chart_status_donut(source, gc, cfg["group"], height=280)
                elif cfg["viz"] in ["Bar Chart", "Column Chart", "Line Chart"]:
                    if cfg["viz"] == "Line Chart":
                        fig = px.line(z, x=gc, y="Count", markers=True, color_discrete_sequence=[GREEN])
                    elif cfg["viz"] == "Bar Chart":
                        fig = px.bar(z, x="Count", y=gc, orientation="h", text="Count", color=gc, color_discrete_sequence=[GREEN, BLUE, ORANGE, RED])
                    else:
                        fig = px.bar(z, x=gc, y="Count", text="Count", color=gc, color_discrete_sequence=[GREEN, BLUE, ORANGE, RED])
                    fig.update_layout(height=280, showlegend=False, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
                elif cfg["viz"] == "Summary Cards":
                    cols = st.columns(min(4, max(1, len(z))))
                    for col, (_, r_row) in zip(cols, z.iterrows()):
                        with col:
                            metric_card(str(r_row[gc]), int(r_row.Count), f"{cfg['type']} records", "📊", "green")

                st.dataframe(source, hide_index=True, use_container_width=True)

                with st.form("save_custom_report_form"):
                    c_a, c_b = st.columns([3, 1])
                    with c_a:
                        report_name = st.text_input("Report Name", f"{cfg['type']} ({cfg['group']}) - {date.today().isoformat()}")
                    with c_b:
                        st.write("")
                        save_rep = st.form_submit_button("＋ Save Current Report", type="primary", use_container_width=True)
                    if save_rep and report_name.strip():
                        sql(
                            "INSERT INTO saved_reports(name,report_type,config,last_generated,created_by) VALUES(?,?,?,?,?)",
                            (report_name.strip(), cfg["type"], json.dumps(cfg), date.today().isoformat(), "Alice Morgan"),
                        )
                        log_event("Saved", "Report", None, report_name.strip())
                        st.success("Custom report saved.")
                        st.rerun()

        with st.container(border=True):
            saved = df("SELECT * FROM saved_reports ORDER BY id DESC")
            section_title("Saved Reports", f"{len(saved)} Saved")
            if saved.empty:
                st.caption("No saved reports yet.")
            else:
                render_rich_table(
                    saved,
                    [
                        ("name", "Report Name", "bold"),
                        ("report_type", "Report Type", "badge"),
                        ("last_generated", "Last Generated", "date"),
                        ("created_by", "Created By", "avatar"),
                    ],
                )


# ===========================================================================
# PAGE 8: SETTINGS
# ===========================================================================
elif page == "Settings":
    render_top_header("Settings", "Manage workspace, users, defaults, notifications, integrations, and system preferences", show_filters=False)
    tabs = st.tabs(
        [
            "Workspace",
            "User Management",
            "Project Settings",
            "Task Settings",
            "Milestones",
            "Notifications",
            "Integrations",
            "System",
        ]
    )

    with tabs[0]:
        with st.container(border=True):
            section_title("Workspace Settings", "General branding and localization")
            logo_file = st.file_uploader("Workspace Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
            if logo_file is not None:
                b64_str = base64.b64encode(logo_file.read()).decode("utf-8")
                put_setting("workspace_logo_b64", b64_str)
                st.success("Logo uploaded and applied to sidebar.")

            with st.form("workspace_settings_form"):
                name = st.text_input("Workspace Name", get_setting("workspace_name", "Vektor"))
                desc = st.text_area("Description", get_setting("workspace_description", "Project management workspace for planning, tracking, and delivery."))
                a, b = st.columns(2)
                with a:
                    t_opts = ["HPE Green", "Blue", "Dark"]
                    cur_t = get_setting("theme", "HPE Green")
                    theme = st.selectbox("Theme", t_opts, index=t_opts.index(cur_t) if cur_t in t_opts else 0)
                with b:
                    l_opts = ["English", "Filipino"]
                    cur_l = get_setting("language", "English")
                    language = st.selectbox("Language", l_opts, index=l_opts.index(cur_l) if cur_l in l_opts else 0)
                a, b = st.columns(2)
                with a:
                    f_opts = ["MMM DD, YYYY", "DD/MM/YYYY", "MM/DD/YYYY"]
                    cur_f = get_setting("date_format", "MMM DD, YYYY")
                    fmt_d = st.selectbox("Date Format", f_opts, index=f_opts.index(cur_f) if cur_f in f_opts else 0)
                with b:
                    p_opts = ["Overview", "Projects", "Task List", "Reports"]
                    cur_p = get_setting("default_landing", "Overview")
                    landing = st.selectbox("Default Landing Page", p_opts, index=p_opts.index(cur_p) if cur_p in p_opts else 0)
                if st.form_submit_button("💾 Save Workspace Settings", type="primary"):
                    for k, v in {
                        "workspace_name": name,
                        "workspace_description": desc,
                        "theme": theme,
                        "language": language,
                        "date_format": fmt_d,
                        "default_landing": landing,
                    }.items():
                        put_setting(k, v)
                    log_event("Updated", "Workspace Settings", None, f"Theme={theme}")
                    st.success("Workspace settings saved.")
                    st.rerun()

    with tabs[1]:
        with st.container(border=True):
            section_title("User Management", f"{len(R)} Registered Users")
            render_rich_table(
                R,
                [
                    ("name", "Name", "avatar"),
                    ("email", "Email", "text"),
                    ("role", "Role", "badge"),
                    ("team", "Team", "text"),
                    ("status", "Status", "badge"),
                ],
                max_rows=12,
            )

        c1, c2 = st.columns(2)
        with c1, st.container(border=True):
            section_title("Add New User")
            with st.form("add_user_form"):
                u_name = st.text_input("Full Name *")
                u_email = st.text_input("Email Address *")
                a, b = st.columns(2)
                with a:
                    u_role = st.selectbox("Role", ["Admin", "Project Manager", "Team Member"])
                with b:
                    u_team = st.selectbox("Team", teams_list)
                if st.form_submit_button("＋ Add User", type="primary"):
                    if u_name.strip():
                        sql(
                            "INSERT INTO resources(name,email,team,role,status,capacity) VALUES(?,?,?,?,?,?)",
                            (u_name.strip(), u_email.strip() or f"{u_name.lower().replace(' ', '.')}@vektor.example", u_team, u_role, "Available", 100),
                        )
                        log_event("Created", "User", None, u_name.strip())
                        st.rerun()
                    else:
                        st.error("Name is required.")
        with c2, st.container(border=True):
            section_title("Update or Remove User")
            if not R.empty:
                uid = st.selectbox("Select User", R.id.tolist(), format_func=lambda x: f"{R.loc[R.id == x, 'name'].iloc[0]} ({R.loc[R.id == x, 'role'].iloc[0]})")
                u_row = R[R.id == uid].iloc[0]
                roles = ["Admin", "Project Manager", "Team Member"]
                states = ["Available", "Partial", "On Leave"]
                new_role = st.selectbox("Role", roles, index=roles.index(u_row.role) if u_row.role in roles else 2, key="upd_u_role")
                new_stat = st.selectbox("Status", states, index=states.index(u_row.status) if u_row.status in states else 0, key="upd_u_stat")
                b_a, b_b = st.columns(2)
                with b_a:
                    if st.button("💾 Update User", type="primary", use_container_width=True):
                        sql("UPDATE resources SET role=?, status=? WHERE id=?", (new_role, new_stat, int(uid)))
                        log_event("Updated", "User", int(uid), u_row["name"])
                        st.rerun()
                with b_b:
                    if st.button("🗑 Remove User", use_container_width=True):
                        sql("DELETE FROM resources WHERE id=?", (int(uid),))
                        log_event("Deleted", "User", int(uid), u_row["name"])
                        st.rerun()

    with tabs[2]:
        with st.container(border=True):
            section_title("Project Defaults & Governance")
            with st.form("project_defaults_form"):
                a, b = st.columns(2)
                with a:
                    p_stats = ["Not Started", "On Track", "At Risk", "Delayed"]
                    cur_ps = get_setting("default_project_status", "On Track")
                    status = st.selectbox("Default Project Status", p_stats, index=p_stats.index(cur_ps) if cur_ps in p_stats else 1)
                with b:
                    t_mpl = ["Standard Project", "Software Delivery", "Marketing Campaign", "Infrastructure"]
                    cur_tm = get_setting("project_template", "Standard Project")
                    template = st.selectbox("Default Project Template", t_mpl, index=t_mpl.index(cur_tm) if cur_tm in t_mpl else 0)
                vals = {
                    k: st.checkbox(label, value=get_setting(k, "true" if default else "false") == "true")
                    for k, label, default in [
                        ("auto_project_id", "Enable project key/ID auto-generation", True),
                        ("require_project_desc", "Require project description", True),
                        ("allow_project_attachments", "Allow project attachments", True),
                        ("project_archiving", "Enable project archiving", False),
                    ]
                }
                if st.form_submit_button("💾 Save Project Defaults", type="primary"):
                    put_setting("default_project_status", status)
                    put_setting("project_template", template)
                    for k, v in vals.items():
                        put_setting(k, str(v).lower())
                    st.success("Project defaults saved.")

    with tabs[3]:
        with st.container(border=True):
            section_title("Task Defaults & Workflow Rules")
            with st.form("task_defaults_form"):
                a, b = st.columns(2)
                with a:
                    status = st.selectbox("Default Task Status", ["Not Started", "In Progress", "Completed"])
                with b:
                    priority = st.selectbox("Default Task Priority", ["High", "Medium", "Low"], index=1)
                vals = {
                    k: st.checkbox(label, value=get_setting(k, "true" if default else "false") == "true")
                    for k, label, default in [
                        ("task_dependencies", "Enable task dependencies", True),
                        ("allow_subtasks", "Allow subtasks", True),
                        ("recurring_tasks", "Enable recurring tasks", False),
                        ("require_assignee", "Require assignee", True),
                        ("task_attachments", "Allow task attachments", True),
                    ]
                }
                if st.form_submit_button("💾 Save Task Defaults", type="primary"):
                    put_setting("default_task_status", status)
                    put_setting("default_task_priority", priority)
                    for k, v in vals.items():
                        put_setting(k, str(v).lower())
                    st.success("Task defaults saved.")

    with tabs[4]:
        c1, c2 = st.columns([1, 1.3])
        with c1, st.container(border=True):
            section_title("Milestone Settings")
            with st.form("milestone_defaults_form"):
                status = st.selectbox("Default Milestone Status", ["Not Started", "In Progress", "Completed"])
                vals = {
                    k: st.checkbox(label, value=get_setting(k, "true") == "true")
                    for k, label in [
                        ("milestones_enabled", "Enable milestones"),
                        ("milestones_require_due", "Require due date"),
                        ("milestone_dependencies", "Allow milestone dependencies"),
                        ("milestone_notifications", "Auto-notify on milestone updates"),
                    ]
                }
                if st.form_submit_button("💾 Save Milestone Defaults", type="primary"):
                    put_setting("default_milestone_status", status)
                    for k, v in vals.items():
                        put_setting(k, str(v).lower())
                    st.success("Milestone defaults saved.")

        with c2, st.container(border=True):
            section_title("Add Milestone")
            with st.form("add_milestone_form"):
                m_name = st.text_input("Milestone Name *")
                a, b = st.columns(2)
                with a:
                    m_proj = st.selectbox("Project", projects_list)
                with b:
                    m_team = st.selectbox("Team", teams_list)
                a, b = st.columns(2)
                with a:
                    m_due = st.date_input("Due Date", date.today() + timedelta(days=14))
                with b:
                    m_stat = st.selectbox("Status", ["Not Started", "In Progress", "Completed", "Overdue"])
                if st.form_submit_button("＋ Add Milestone", type="primary"):
                    if m_name.strip():
                        sql(
                            "INSERT INTO milestones(name,project_id,team,status,start_date,due_date,notes) VALUES(?,?,?,?,?,?,?)",
                            (m_name.strip(), pmap[m_proj], m_team, m_stat, date.today().isoformat(), m_due.isoformat(), ""),
                        )
                        log_event("Created", "Milestone", None, m_name.strip())
                        st.rerun()

        with st.container(border=True):
            section_title("Active Milestones", f"{len(M)} Total")
            render_rich_table(
                M,
                [
                    ("name", "Milestone", "bold"),
                    ("project", "Project", "text"),
                    ("team", "Team", "text"),
                    ("due_date", "Due Date", "date"),
                    ("status", "Status", "badge"),
                ],
            )

    with tabs[5]:
        with st.container(border=True):
            section_title("Notification Preferences", "Configure triggers and delivery cadence")
            with st.form("notifications_form"):
                labels = [
                    "Task assigned to me",
                    "Task due soon",
                    "Task overdue",
                    "Milestone updates",
                    "Project status changes",
                    "Project assignment",
                    "Mention in comments",
                    "Weekly project digest",
                ]
                c_a, c_b = st.columns(2)
                vals = {}
                for idx, k in enumerate(labels):
                    target_col = c_a if idx < 4 else c_b
                    with target_col:
                        vals[k] = st.checkbox(k, value=get_setting("notify_" + k, "false" if k == "Mention in comments" else "true") == "true")
                freq_opts = ["Real-time", "Hourly", "Daily digest", "Weekly digest"]
                cur_freq = get_setting("notification_frequency", "Daily digest")
                freq = st.selectbox("Notification Frequency", freq_opts, index=freq_opts.index(cur_freq) if cur_freq in freq_opts else 2)
                if st.form_submit_button("💾 Save Notification Preferences", type="primary"):
                    for k, v in vals.items():
                        put_setting("notify_" + k, str(v).lower())
                    put_setting("notification_frequency", freq)
                    st.success("Notification preferences saved.")

    with tabs[6]:
        section_title("Connected Apps & Integrations", "Sync calendars, messaging channels, and cloud storage")
        services = [
            ("Microsoft 365", "Connect Outlook calendar, email alerts, and directory sync", "📅"),
            ("Slack", "Send automated task, milestone, and health alerts to Slack channels", "💬"),
            ("Microsoft Teams", "Collaborate on project channels and receive @mention notifications", "👥"),
            ("OneDrive", "Attach and sync project deliverables and specification documents", "☁️"),
        ]
        cols = st.columns(2)
        for idx, (service, desc, ico) in enumerate(services):
            with cols[idx % 2], st.container(border=True):
                is_conn = get_setting(f"integration_{service}", "false") == "true"
                a, b = st.columns([3.2, 1.1])
                with a:
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:10px'>"
                        f"<span style='font-size:24px'>{ico}</span>"
                        f"<div><div style='font-weight:700;font-size:15px'>{html.escape(service)}</div>"
                        f"<div style='font-size:12px;color:#64748B'>{html.escape(desc)}</div></div></div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        f"<div style='margin-top:8px'>{badge_html('Connected' if is_conn else 'Disconnected')}</div>",
                        unsafe_allow_html=True,
                    )
                with b:
                    st.write("")
                    btn_label = "Disconnect" if is_conn else "Connect"
                    if st.button(btn_label, key=f"int_{service}", type="secondary" if is_conn else "primary", use_container_width=True):
                        put_setting(f"integration_{service}", "false" if is_conn else "true")
                        log_event("Integration Toggled", service, None, "Disconnected" if is_conn else "Connected")
                        st.rerun()

    with tabs[7]:
        c1, c2 = st.columns([1.2, 1])
        with c1, st.container(border=True):
            section_title("System & Security Preferences")
            with st.form("system_settings_form"):
                audit = st.checkbox("Enable audit logs", value=get_setting("audit_logs", "true") == "true")
                exports = st.checkbox("Allow data export", value=get_setting("data_export", "true") == "true")
                backups = st.checkbox("Enable automatic backups", value=get_setting("automatic_backups", "true") == "true")
                maintenance = st.checkbox("Enable maintenance mode", value=get_setting("maintenance_mode", "false") == "true")
                t_opts = ["30 minutes", "1 hour", "4 hours", "8 hours", "12 hours"]
                cur_to = get_setting("session_timeout", "8 hours")
                timeout = st.selectbox("Session Timeout", t_opts, index=t_opts.index(cur_to) if cur_to in t_opts else 3)
                if st.form_submit_button("💾 Save System Settings", type="primary"):
                    for k, v in {
                        "audit_logs": audit,
                        "data_export": exports,
                        "automatic_backups": backups,
                        "maintenance_mode": maintenance,
                        "session_timeout": timeout,
                    }.items():
                        put_setting(k, str(v).lower() if isinstance(v, bool) else v)
                    st.success("System settings saved.")

        with c2, st.container(border=True):
            section_title("Database & Demo Data Management")
            if os.path.exists(DB):
                with open(DB, "rb") as f:
                    db_bytes = f.read()
                st.download_button(
                    "⬇ Download SQLite Backup (.db)",
                    db_bytes,
                    file_name=f"vektor_backup_{date.today().isoformat()}.db",
                    mime="application/octet-stream",
                    type="primary",
                    use_container_width=True,
                )
            if st.button("🔄 Reset Workspace to Default Demo Data", use_container_width=True):
                init_db(force_reset=True)
                st.success("Database reset to factory demo data.")
                st.rerun()
            if st.button("🧹 Clear Audit Log", use_container_width=True):
                sql("DELETE FROM audit_log")
                st.success("Audit log cleared.")
                st.rerun()

        with st.container(border=True):
            logs = df("SELECT * FROM audit_log ORDER BY id DESC LIMIT 50")
            section_title("Recent Audit Events", f"{len(logs)} Logged")
            if logs.empty:
                st.caption("No audit events recorded yet.")
            else:
                render_rich_table(
                    logs,
                    [
                        ("created_at", "Timestamp", "text"),
                        ("event", "Action", "badge"),
                        ("item_type", "Entity", "bold"),
                        ("details", "Details", "text"),
                    ],
                )
