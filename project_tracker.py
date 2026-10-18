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
# Database Layer & Seeding
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
        if force_reset:
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
                status TEXT,
                start_date TEXT,
                end_date TEXT,
                progress INTEGER,
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
            ps = [
                ("Website Revamp", "James Miller", "IT", "On Track", "2024-04-01", "2024-09-30", 90, "Complete redesign and migration of corporate web properties.", "Healthy"),
                ("Marketing Campaign", "Sarah Kim", "Marketing", "At Risk", "2024-04-15", "2024-08-31", 60, "Q2/Q3 integrated multi-channel product launch campaign.", "At Risk"),
                ("IT Infrastructure Upgrade", "Michael Tan", "IT", "At Risk", "2024-05-01", "2024-10-31", 40, "Core network, server virtualization, and cloud security upgrade.", "At Risk"),
                ("Sustainability Initiative", "Alice Morgan", "Operations", "Delayed", "2024-05-15", "2024-12-15", 25, "Carbon footprint tracking and green supply chain initiative.", "Critical"),
                ("Internal Operations", "Ryan Cruz", "Operations", "On Track", "2024-06-01", "2024-11-30", 75, "Workflow automation and SOP standardization across departments.", "Healthy"),
            ]
            c.executemany(
                "INSERT INTO projects(name,owner,team,status,start_date,end_date,progress,description,health) VALUES(?,?,?,?,?,?,?,?,?)",
                ps,
            )

            people = [
                "James Miller", "Sarah Kim", "Alice Morgan", "Ryan Cruz", "Michael Tan",
                "Daniel Perez", "Ella Ramirez", "Chris Padilla", "Anna Reyes", "Kevin Santos",
                "Lisa Chan", "David Lim", "Maria Santos", "John Cruz", "Paolo Reyes",
                "Nina Flores", "Mark Lee", "Jade Tan", "Ben Garcia", "Mia Lopez",
                "Leo Ramos", "Ivy Chen", "Noah King", "Zoe Park",
            ]
            teams_list = [
                "IT", "Design", "Marketing", "Operations", "IT",
                "Operations", "Design", "IT", "Marketing", "Operations",
                "Design", "IT", "Marketing", "Operations", "IT",
                "Design", "Marketing", "Operations", "IT", "Design",
                "IT", "Marketing", "Operations", "Others",
            ]
            for i, (n, t) in enumerate(zip(people, teams_list)):
                status = "On Leave" if i in (4, 10) else ("Partial" if i in (2, 6, 14, 17) else "Available")
                role = "Admin" if i == 0 else ("Project Manager" if i in (1, 2, 3, 4) else "Team Member")
                cap = 0 if status == "On Leave" else (50 if status == "Partial" else 100)
                c.execute(
                    "INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",
                    (n, n.lower().replace(" ", ".") + "@vektor.example", t, role, status, cap, ""),
                )

            tasks_seed = [
                ("UI/UX Design Complete", 1, "Design", "Completed", "High", "2024-04-01", "2024-04-30", "Alice Morgan", 100, "", "Design phase completed."),
                ("Homepage Development", 1, "IT", "In Progress", "High", "2024-04-10", "2024-05-15", "James Miller", 65, "UI/UX Design Complete", "Frontend responsive build."),
                ("Content Review", 2, "Marketing", "In Progress", "Medium", "2024-04-15", "2024-05-20", "Sarah Kim", 55, "", "Copywriting and creative review."),
                ("Infrastructure Setup", 3, "IT", "Not Started", "High", "2024-05-01", "2024-05-15", "Michael Tan", 0, "Requirements Gathering", "Provisioning cluster nodes."),
                ("Sustainability Plan", 4, "Operations", "Completed", "Medium", "2024-05-05", "2024-05-30", "Ryan Cruz", 100, "", "Charter approved by leadership."),
                ("Process Documentation", 5, "Operations", "In Progress", "Low", "2024-05-10", "2024-06-20", "Daniel Perez", 50, "", "Mapping core operational workflows."),
                ("Training Materials", 2, "Marketing", "Not Started", "Medium", "2024-05-20", "2024-06-30", "Ella Ramirez", 0, "Content Review", "Sales enablement deck."),
                ("QA Testing", 1, "IT", "In Progress", "High", "2024-05-25", "2024-06-25", "Chris Padilla", 35, "Homepage Development", "Cross-browser regression testing."),
                ("Campaign Launch", 2, "Marketing", "Not Started", "High", "2024-05-25", "2024-06-10", "Sarah Kim", 0, "Content Review", "Go-live across paid channels."),
                ("System Configuration", 3, "IT", "Completed", "Low", "2024-03-15", "2024-04-10", "Michael Tan", 100, "", "Baseline firewall and VLAN setup."),
                ("Requirements Gathering", 5, "Operations", "Completed", "Low", "2024-04-01", "2024-04-25", "Daniel Perez", 100, "", "Stakeholder interviews finished."),
            ]
            c.executemany(
                "INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                tasks_seed,
            )

            ms_seed = [
                ("UI/UX Design Complete", 1, "Design", "Completed", "2024-04-01", "2024-04-30", "2024-04-28", "", "Design phase completed."),
                ("Campaign Launch", 2, "Marketing", "In Progress", "2024-04-15", "2024-05-20", "", "Content Review", "Awaiting final assets."),
                ("Infrastructure Setup", 3, "IT", "In Progress", "2024-05-01", "2024-05-15", "", "Requirements Gathering", "Environment setup ongoing."),
                ("Sustainability Plan", 4, "Operations", "Not Started", "2024-05-05", "2024-05-25", "", "", "Pending approval."),
                ("Process Documentation", 5, "Operations", "Not Started", "2024-05-10", "2024-05-30", "", "Infrastructure Setup", ""),
                ("QA Testing", 1, "IT", "Overdue", "2024-04-01", "2024-04-20", "", "Development Complete", "Testing issues to resolve."),
                ("Training Materials", 2, "Marketing", "Overdue", "2024-04-20", "2024-05-05", "", "Content Review", ""),
                ("System Configuration", 3, "IT", "Completed", "2024-03-15", "2024-04-10", "2024-04-09", "", "Configuration completed."),
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
                    ("Monthly Project Summary", "Project Report", json.dumps({"type": "Project Report", "group": "Status", "metrics": ["Project Status", "Progress"], "viz": "Bar Chart"}), "2024-04-20", "James Miller"),
                    ("Q2 Resource Allocation", "Resource Report", json.dumps({"type": "Resource Report", "group": "Team", "metrics": ["Team Members"], "viz": "Column Chart"}), "2024-04-18", "Sarah Kim"),
                    ("Overdue Tasks & Milestones", "Task Report", json.dumps({"type": "Task Report", "group": "Priority", "metrics": ["Task Status", "Priority"], "viz": "Pie Chart"}), "2024-04-15", "Michael Tan"),
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
    "linear-gradient(180deg, #001F1D 0%, #002F2A 100%)"
    if active_theme == "HPE Green"
    else ("linear-gradient(180deg, #0B1938 0%, #132A5E 100%)" if active_theme == "Blue" else "linear-gradient(180deg, #0B0F19 0%, #151C2C 100%)")
)
APP_BG = (
    "linear-gradient(135deg, #F3FAFF 0%, #EAF4FC 100%)"
    if active_theme != "Dark"
    else "linear-gradient(135deg, #0F172A 0%, #1E293B 100%)"
)
CARD_BG = "#FFFFFF" if active_theme != "Dark" else "#1E293B"
TEXT_COLOR = NAVY if active_theme != "Dark" else "#F8FAFC"
BORDER_COLOR = "#DCE8F3" if active_theme != "Dark" else "#334155"


# ---------------------------------------------------------------------------
# Global CSS Styling (Exact Replica of Vektor Mockups)
# ---------------------------------------------------------------------------
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: {TEXT_COLOR};
}}
.stApp {{
    background: {APP_BG};
    color: {TEXT_COLOR};
}}
.block-container {{
    padding-top: 1.1rem !important;
    padding-bottom: 2rem !important;
    max-width: 1820px;
}}
header[data-testid="stHeader"] {{
    background: transparent !important;
}}

/* Sidebar Styling */
[data-testid="stSidebar"] {{
    background: {SIDEBAR_BG} !important;
    border-right: 1px solid rgba(255,255,255,0.08);
}}
[data-testid="stSidebar"] * {{
    color: #F4FBFA !important;
}}
[data-testid="stSidebar"] .stButton > button {{
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 0.6rem 0.9rem !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    border: 1px solid transparent !important;
    background: transparent !important;
    transition: all 0.15s ease;
}}
[data-testid="stSidebar"] .stButton > button:hover {{
    background: rgba(255, 255, 255, 0.08) !important;
}}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
    background: {PRIMARY_ACCENT} !important;
    color: #FFFFFF !important;
    box-shadow: 0 4px 12px rgba(0, 135, 90, 0.35);
    border-left: 4px solid #6EE7B7 !important;
}}

/* Primary & Secondary Buttons in Main Area */
.stButton > button, .stDownloadButton > button {{
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    border: 1px solid {BORDER_COLOR} !important;
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
    padding: 14px !important;
    box-shadow: 0 2px 10px rgba(16, 25, 79, 0.04);
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

/* Custom Metric Cards */
.vk-metric-card {{
    display: flex;
    align-items: center;
    gap: 14px;
    background: {CARD_BG};
    border: 1px solid {BORDER_COLOR};
    border-radius: 12px;
    padding: 15px 18px;
    min-height: 94px;
    box-shadow: 0 2px 10px rgba(16, 25, 79, 0.04);
}}
.vk-metric-icon {{
    width: 50px;
    height: 50px;
    min-width: 50px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 50%;
    font-size: 22px;
    font-weight: 700;
}}
.vk-metric-label {{
    font-size: 12.5px;
    font-weight: 600;
    color: #5A6A85;
    margin-bottom: 2px;
}}
.vk-metric-value {{
    font-size: 28px;
    font-weight: 800;
    color: {TEXT_COLOR};
    line-height: 1.15;
}}
.vk-metric-note {{
    font-size: 11.5px;
    font-weight: 500;
    color: #64748B;
    margin-top: 2px;
}}

/* Section Headers */
.vk-section-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 15.5px;
    font-weight: 700;
    color: {TEXT_COLOR};
    margin-bottom: 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid {BORDER_COLOR};
}}

/* Custom HTML Table Styling */
.vk-table-wrap {{
    overflow-x: auto;
    border-radius: 10px;
    border: 1px solid {BORDER_COLOR};
    background: {CARD_BG};
}}
table.vk-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    text-align: left;
}}
table.vk-table thead tr {{
    background: #F4F8FB;
    border-bottom: 1px solid {BORDER_COLOR};
}}
table.vk-table th {{
    padding: 10px 12px;
    font-weight: 700;
    font-size: 12px;
    color: #475569;
    text-transform: uppercase;
    letter-spacing: 0.02em;
}}
table.vk-table td {{
    padding: 10px 12px;
    border-bottom: 1px solid #EEF4F8;
    color: {TEXT_COLOR};
    vertical-align: middle;
}}
table.vk-table tbody tr:hover {{
    background: #F8FBFF;
}}

/* Status Badges / Pills */
.vk-badge {{
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 11.5px;
    font-weight: 600;
    white-space: nowrap;
}}
.vk-badge-green {{ background: #E6F6F0; color: #007A50; border: 1px solid #B7EBD6; }}
.vk-badge-blue {{ background: #E8F2FF; color: #0F62FE; border: 1px solid #BBD6FF; }}
.vk-badge-orange {{ background: #FFF4E5; color: #D97706; border: 1px solid #FCD39B; }}
.vk-badge-red {{ background: #FEECEB; color: #DC2626; border: 1px solid #FCA5A5; }}
.vk-badge-gray {{ background: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; }}
.vk-badge-purple {{ background: #F3E8FF; color: #7C3AED; border: 1px solid #DDD6FE; }}

/* Progress Bar inside Tables */
.vk-prog-wrap {{
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 115px;
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
    font-weight: 700;
    min-width: 34px;
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
    font-size: 10.5px;
    font-weight: 700;
    color: white;
    margin-right: 7px;
}}

/* Top Bar User & Breadcrumb */
.vk-topbar {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: {CARD_BG};
    border: 1px solid {BORDER_COLOR};
    border-radius: 12px;
    padding: 10px 18px;
    margin-bottom: 14px;
    box-shadow: 0 2px 8px rgba(16, 25, 79, 0.03);
}}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# UI Helper Functions
# ---------------------------------------------------------------------------
def badge_html(val: object) -> str:
    s = str(val or "Unknown")
    mapping = {
        "On Track": ("vk-badge-green", "●"),
        "Completed": ("vk-badge-green", "✓"),
        "Healthy": ("vk-badge-green", "●"),
        "Available": ("vk-badge-green", "●"),
        "Balanced": ("vk-badge-green", "●"),
        "Connected": ("vk-badge-green", "●"),
        "Active": ("vk-badge-green", "●"),
        "Low": ("vk-badge-green", "↓"),
        "In Progress": ("vk-badge-blue", "◐"),
        "Admin": ("vk-badge-purple", "★"),
        "Project Manager": ("vk-badge-blue", "◆"),
        "Team Member": ("vk-badge-gray", "•"),
        "At Risk": ("vk-badge-orange", "▲"),
        "Partial": ("vk-badge-orange", "◐"),
        "Medium": ("vk-badge-orange", "•"),
        "High": ("vk-badge-red", "▲"),
        "Delayed": ("vk-badge-red", "●"),
        "Critical": ("vk-badge-red", "⚠"),
        "Overdue": ("vk-badge-red", "⚠"),
        "On Leave": ("vk-badge-red", "●"),
        "Overloaded": ("vk-badge-red", "⚠"),
        "Not Started": ("vk-badge-gray", "○"),
        "Disconnected": ("vk-badge-gray", "○"),
    }
    cls, dot = mapping.get(s, ("vk-badge-gray", "•"))
    return f'<span class="vk-badge {cls}">{dot} {html.escape(s)}</span>'


def progress_html(val: object) -> str:
    try:
        pct = max(0, min(100, int(float(val or 0))))
    except Exception:
        pct = 0
    color = GREEN if pct >= 75 else (BLUE if pct >= 50 else (ORANGE if pct >= 30 else RED))
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
    palette = [GREEN, BLUE, ORANGE, PURPLE, "#0EA5E9", "#EC4899", "#14B8A6"]
    bg = palette[sum(ord(ch) for ch in s) % len(palette)]
    return f'<span class="vk-avatar" style="background:{bg}">{html.escape(initials)}</span>{html.escape(s)}'


def fmt_date(d_str: object) -> str:
    if not d_str or pd.isna(d_str) or not str(d_str).strip():
        return "—"
    try:
        d = date.fromisoformat(str(d_str)[:10])
        fmt_pref = get_setting("date_format", "MMM DD, YYYY")
        if fmt_pref == "DD/MM/YYYY":
            return d.strftime("%d/%m/%Y")
        if fmt_pref == "MM/DD/YYYY":
            return d.strftime("%m/%d/%Y")
        return d.strftime("%b %d, %Y")
    except Exception:
        return str(d_str)


def render_rich_table(
    data: pd.DataFrame,
    columns: list[tuple[str, str, str]],
    max_rows: int | None = None,
) -> None:
    """Render a custom HTML table matching Vektor mockup styling.
    columns: list of (df_col, header_label, col_type)
    col_type in {'text', 'bold', 'badge', 'progress', 'avatar', 'date'}
    """
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
                cell = f'<span style="font-weight:650;color:{TEXT_COLOR}">{html.escape(str(val if pd.notna(val) else "—"))}</span>'
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
        "Low": GREEN,
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
    fig = px.pie(
        z,
        names="Status",
        values="Count",
        hole=0.62,
        color="Status",
        color_discrete_map=colors,
    )
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
    bar_colors = [
        GREEN if v >= 75 else (BLUE if v >= 50 else (ORANGE if v >= 30 else RED))
        for v in z[val]
    ]
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
    ws_name = get_setting("workspace_name", "Vektor")
    overdue_cnt = int((M.status == "Overdue").sum() + (T.status == "Overdue").sum())
    st.markdown(
        f"""
        <div class="vk-topbar">
            <div style="display:flex;align-items:center;gap:10px">
                <span style="font-weight:800;color:{PRIMARY_ACCENT};font-size:14px">{html.escape(ws_name)}</span>
                <span style="color:#94A3B8">/</span>
                <span style="font-weight:700;font-size:14px;color:{TEXT_COLOR}">{html.escape(title)}</span>
                <span style="font-size:12px;color:#64748B;margin-left:8px">• {html.escape(subtitle)}</span>
            </div>
            <div style="display:flex;align-items:center;gap:14px">
                <span class="vk-badge {'vk-badge-red' if overdue_cnt else 'vk-badge-green'}">
                    🔔 {overdue_cnt} Overdue Alerts
                </span>
                <div style="display:flex;align-items:center;gap:8px;font-size:13px;font-weight:600">
                    <span class="vk-avatar" style="background:{PRIMARY_ACCENT}">JM</span>
                    <span>James Miller <span style="color:#64748B;font-weight:400">(Admin)</span></span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if show_filters:
        c1, c2, c3, c4 = st.columns([1.7, 1, 1, 0.9])
        with c1:
            st.markdown(
                f"<div style='font-size:26px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.1'>{html.escape(title)}</div>"
                f"<div style='font-size:13px;color:#64748B;margin-top:2px'>{html.escape(subtitle)}</div>",
                unsafe_allow_html=True,
            )
        with c2:
            st.selectbox("Project", ["All Projects"] + projects_list, key="filter_project")
        with c3:
            st.selectbox("Team", ["All Teams"] + teams_list, key="filter_team")
        with c4:
            st.write("")
            st.download_button(
                "⬇ Export Portfolio",
                excel_bytes({"Projects": P, "Tasks": T, "Resources": R, "Milestones": M}),
                file_name="vektor_portfolio.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.markdown(
            f"<div style='font-size:26px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.1;margin-bottom:4px'>{html.escape(title)}</div>"
            f"<div style='font-size:13px;color:#64748B;margin-bottom:10px'>{html.escape(subtitle)}</div>",
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
# Sidebar Navigation
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
        <div style="padding: 4px 4px 18px 4px;">
            <div style="font-size:30px;font-weight:800;letter-spacing:-1.2px;line-height:1;">
                {html.escape(ws_title)} <span style="color:{TEAL_ACCENT}">❯❯</span>
            </div>
            <div style="font-size:11.5px;color:#A7F3D0 !important;opacity:0.85;margin-top:4px;font-weight:500;">
                Direction + Velocity
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    nav_items = {
        "▦  Overview": "Overview",
        "▤  Projects": "Projects",
        "📊  Gantt Chart": "Gantt Chart",
        "☷  Task List": "Task List",
        "👥  Resources": "Resources",
        "📅  Calendar": "Calendar",
        "📈  Reports": "Reports",
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

    st.markdown("<hr style='border-color:rgba(255,255,255,0.12);margin:18px 0'>", unsafe_allow_html=True)
    st.caption("PROJECT WORKSPACE")
    st.markdown(f"**{html.escape(ws_title)} Workspace**")
    st.caption(f"{len(P)} Active Projects • {len(R)} Members")

    if st.button("＋ New Project", use_container_width=True, type="primary", key="sidebar_new_proj"):
        st.session_state.page = "Projects"
        st.session_state.projects_tab_idx = 1
        st.rerun()

page = st.session_state.page


# ===========================================================================
# PAGE 1: OVERVIEW
# ===========================================================================
if page == "Overview":
    render_top_header("Overview", "Your project portfolio, task velocity, and team workload at a glance", show_filters=True)
    p = apply_global_filters(P, "name")
    t = apply_global_filters(T, "project")
    m = apply_global_filters(M, "project")
    r = apply_global_filters(R, "none")

    # KPI Row
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("Total Projects", len(p), f"{int((p.status == 'On Track').sum())} on track • {int(p.status.isin(['At Risk', 'Delayed']).sum())} at risk/delayed", "📁", "green")
    with k2:
        metric_card("Total Tasks", len(t), f"{int((t.status == 'Completed').sum())} completed • {int((t.status == 'In Progress').sum())} in progress", "☑️", "blue")
    with k3:
        metric_card("Team Members", len(r), f"{int((r.status == 'Available').sum())} available • {int((r.status == 'Partial').sum())} partial", "👥", "orange")
    with k4:
        avg_prog = int(round(p.progress.mean())) if len(p) else 0
        metric_card("Overall Progress", f"{avg_prog}%", "Portfolio weighted average", "◉", "purple")

    st.write("")
    # Row 1: 3 Status Donut Charts
    c1, c2, c3 = st.columns(3)
    with c1, st.container(border=True):
        section_title("Project Status", f"{len(p)} Total")
        chart_status_donut(p, "status", "Projects")
    with c2, st.container(border=True):
        section_title("Task Status", f"{len(t)} Total")
        chart_status_donut(t, "status", "Tasks")
    with c3, st.container(border=True):
        section_title("Team Availability", f"{len(r)} Members")
        chart_status_donut(r, "status", "Members")

    # Row 2: Project Progress, Task Completion Trend, Team Workload
    c1, c2, c3 = st.columns(3)
    with c1, st.container(border=True):
        section_title("Project Progress", "Completion %")
        progress_bar_chart(p.sort_values("progress", ascending=True), "name", "progress", height=250)
    with c2, st.container(border=True):
        section_title("Task Completion Trend", "Apr – Sep 2024")
        trend = pd.DataFrame(
            {
                "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                "Completed": [8, 18, 27, 35, 40, 48],
                "In Progress": [3, 12, 16, 19, 21, 24],
                "Not Started": [10, 8, 7, 5, 4, 2],
            }
        )
        fig = px.line(
            trend,
            x="Month",
            y=["Completed", "In Progress", "Not Started"],
            markers=True,
            color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": GRAY},
        )
        fig.update_layout(
            height=250,
            margin=dict(l=0, r=0, t=6, b=0),
            xaxis_title=None,
            yaxis_title=None,
            legend_title_text="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    with c3, st.container(border=True):
        section_title("Team Workload", "Active Tasks by Assignee")
        cnt = t[t.assignee.fillna("") != ""].groupby("assignee").size().rename("Tasks").reset_index().sort_values("Tasks", ascending=True).tail(6)
        if not cnt.empty:
            fig = px.bar(
                cnt,
                x="Tasks",
                y="assignee",
                orientation="h",
                text="Tasks",
                color="Tasks",
                color_continuous_scale=[GREEN, BLUE, ORANGE],
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(
                height=250,
                margin=dict(l=0, r=28, t=6, b=0),
                xaxis_title=None,
                yaxis_title=None,
                coloraxis_showscale=False,
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("No assigned tasks.")

    # Row 3: Upcoming Milestones, Top Projects by Progress, Project Health
    c1, c2, c3 = st.columns([1.35, 1, 1.05])
    with c1, st.container(border=True):
        section_title("Upcoming Milestones", "Next 6 Deadlines")
        render_rich_table(
            m,
            [
                ("name", "Milestone", "bold"),
                ("project", "Project", "text"),
                ("due_date", "Due Date", "date"),
                ("status", "Status", "badge"),
            ],
            max_rows=6,
        )
    with c2, st.container(border=True):
        section_title("Top Projects by Progress", "Leaderboards")
        render_rich_table(
            p.sort_values("progress", ascending=False),
            [
                ("name", "Project", "bold"),
                ("owner", "Owner", "avatar"),
                ("progress", "Progress", "progress"),
            ],
            max_rows=5,
        )
    with c3, st.container(border=True):
        section_title("Project Health", "Risk & Delivery Status")
        render_rich_table(
            p,
            [
                ("name", "Project", "bold"),
                ("status", "Status", "badge"),
                ("progress", "Progress", "progress"),
                ("health", "Health", "badge"),
            ],
            max_rows=5,
        )


# ===========================================================================
# PAGE 2: PROJECTS
# ===========================================================================
elif page == "Projects":
    render_top_header("Projects", "Plan, organize, and monitor your project portfolio", show_filters=True)
    p = apply_global_filters(P, "name")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("Total Projects", len(p), "Active portfolio", "📁", "green")
    with k2:
        metric_card("On Track", int((p.status == "On Track").sum()), "Delivering on schedule", "✅", "green")
    with k3:
        metric_card("At Risk", int((p.status == "At Risk").sum()), "Requires attention", "⚠️", "orange")
    with k4:
        metric_card("Delayed", int((p.status == "Delayed").sum()), "Behind schedule", "🔴", "red")

    st.write("")
    tabs = st.tabs(["▤ Project Portfolio", "＋ Create Project", "✎ Edit / Manage Project"])

    with tabs[0]:
        with st.container(border=True):
            f1, f2, f3 = st.columns([2, 1, 1])
            with f1:
                q = st.text_input("🔍 Search projects by name, owner, or description", key="proj_search")
            with f2:
                sf = st.selectbox("Status Filter", ["All Statuses", "On Track", "At Risk", "Delayed", "Not Started"])
            with f3:
                hf = st.selectbox("Health Filter", ["All Health", "Healthy", "At Risk", "Critical"])

            v = p.copy()
            if q:
                v = v[v.astype(str).apply(lambda s: s.str.contains(q, case=False, na=False)).any(axis=1)]
            if sf != "All Statuses":
                v = v[v.status == sf]
            if hf != "All Health":
                v = v[v.health == hf]

            section_title("Project Portfolio Directory", f"{len(v)} Projects")
            render_rich_table(
                v,
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
            st.write("")
            st.download_button(
                "⬇ Export Projects CSV",
                v.to_csv(index=False).encode("utf-8-sig"),
                file_name="vektor_projects.csv",
                mime="text/csv",
            )

    with tabs[1]:
        with st.container(border=True):
            section_title("Create New Project", "Add a project to the Vektor portfolio")
            with st.form("new_project_form"):
                name = st.text_input("Project Name *")
                a, b, c = st.columns(3)
                with a:
                    owner = st.selectbox("Project Owner", R.name.tolist() if not R.empty else ["James Miller"])
                with b:
                    team = st.selectbox("Team", teams_list)
                with c:
                    def_stat = get_setting("default_project_status", "On Track")
                    stat_opts = ["Not Started", "On Track", "At Risk", "Delayed"]
                    status = st.selectbox("Status", stat_opts, index=stat_opts.index(def_stat) if def_stat in stat_opts else 1)
                a, b, c = st.columns(3)
                with a:
                    start = st.date_input("Start Date", date.today())
                with b:
                    end = st.date_input("End Date", date.today() + timedelta(days=90))
                with c:
                    prog = st.slider("Initial Progress (%)", 0, 100, 0)
                desc = st.text_area("Project Description")
                save = st.form_submit_button("＋ Create Project", type="primary")
                if save:
                    req_desc = get_setting("require_project_desc", "true") == "true"
                    if not name.strip() or end < start:
                        st.error("Please provide a valid project name and date range.")
                    elif req_desc and not desc.strip():
                        st.error("Project description is required by workspace settings.")
                    else:
                        health = "Critical" if status == "Delayed" else ("At Risk" if status == "At Risk" else "Healthy")
                        sql(
                            "INSERT INTO projects(name,owner,team,status,start_date,end_date,progress,description,health) VALUES(?,?,?,?,?,?,?,?,?)",
                            (name.strip(), owner, team, status, start.isoformat(), end.isoformat(), prog, desc, health),
                        )
                        log_event("Created", "Project", None, name.strip())
                        st.success(f"Project '{name}' created.")
                        st.rerun()

    with tabs[2]:
        if P.empty:
            st.info("No projects available to edit.")
        else:
            with st.container(border=True):
                opts = {f"{r['name']} (#{r['id']})": int(r.id) for _, r in P.iterrows()}
                chosen = st.selectbox("Select Project to Edit", list(opts.keys()))
                pid = opts[chosen]
                row = P[P.id == pid].iloc[0]
                with st.form("edit_project_form"):
                    name = st.text_input("Project Name", row["name"])
                    a, b, c = st.columns(3)
                    with a:
                        owner = st.text_input("Owner", row.owner or "")
                    with b:
                        team = st.selectbox("Team", teams_list, index=teams_list.index(row.team) if row.team in teams_list else 0)
                    with c:
                        stat_opts = ["Not Started", "On Track", "At Risk", "Delayed"]
                        status = st.selectbox("Status", stat_opts, index=stat_opts.index(row.status) if row.status in stat_opts else 0)
                    a, b, c = st.columns(3)
                    with a:
                        start = st.date_input("Start Date", date.fromisoformat(row.start_date))
                    with b:
                        end = st.date_input("End Date", date.fromisoformat(row.end_date))
                    with c:
                        progress = st.slider("Progress (%)", 0, 100, int(row.progress))
                    desc = st.text_area("Description", row.description or "")
                    save = st.form_submit_button("💾 Save Project Changes", type="primary")
                    if save:
                        if not name.strip() or end < start:
                            st.error("Enter a valid name and date range.")
                        else:
                            health = "Critical" if status == "Delayed" else ("At Risk" if status == "At Risk" or progress < 50 else "Healthy")
                            sql(
                                "UPDATE projects SET name=?,owner=?,team=?,status=?,start_date=?,end_date=?,progress=?,description=?,health=? WHERE id=?",
                                (name.strip(), owner, team, status, start.isoformat(), end.isoformat(), progress, desc, health, pid),
                            )
                            log_event("Updated", "Project", pid, name.strip())
                            st.success("Project updated.")
                            st.rerun()

                with st.expander("🗑 Danger Zone: Delete Project"):
                    st.warning("Deleting a project will also remove its linked tasks and milestones.")
                    if st.button("Delete Selected Project", key="del_proj_btn"):
                        sql("DELETE FROM tasks WHERE project_id=?", (pid,))
                        sql("DELETE FROM milestones WHERE project_id=?", (pid,))
                        sql("DELETE FROM projects WHERE id=?", (pid,))
                        log_event("Deleted", "Project", pid, chosen)
                        st.rerun()


# ===========================================================================
# PAGE 3: GANTT CHART
# ===========================================================================
elif page == "Gantt Chart":
    render_top_header("Gantt Chart", "Visualize project schedules, task durations, dependencies, and milestones", show_filters=True)
    t = apply_global_filters(T, "project")
    m = apply_global_filters(M, "project")

    with st.container(border=True):
        c1, c2, c3 = st.columns([1.2, 1, 1])
        with c1:
            view_mode = st.radio("Timeline View", ["Tasks & Milestones", "Tasks Only", "Projects Timeline"], horizontal=True)
        with c2:
            color_by = st.selectbox("Color Bars By", ["status", "priority", "project", "team"])
        with c3:
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
            mode = st.radio("Action", ["Create Task", "Edit Existing Task"], horizontal=True)
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
            mode = st.radio("Resource Action", ["Add Resource", "Edit Resource"], horizontal=True)
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
        c1, c2 = st.columns([1, 2.5])
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
# PAGE 7: REPORTS (Exact Replica of All 7 Mockup Report Tabs)
# ===========================================================================
elif page == "Reports":
    render_top_header("Reports", "Gain insights into project progress, resource utilization, and team performance.", show_filters=False)

    # Mockup Filter Bar: Date Range | Project | Team | Export Format | Export Report
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

    # --- REPORT TAB 0: OVERVIEW ---
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

    # --- REPORT TAB 1: PROJECTS ---
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

    # --- REPORT TAB 2: RESOURCES ---
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

    # --- REPORT TAB 3: TASKS ---
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

    # --- REPORT TAB 4: MILESTONES ---
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

    # --- REPORT TAB 5: WORKLOAD ---
    with tabs[5]:
        counts = T.groupby("assignee").size().rename("active_tasks").reset_index()
        work = rr.merge(counts, left_on="name", right_on="assignee", how="left").drop(columns=["assignee"], errors="ignore")
        work["active_tasks"] = work["active_tasks"].fillna(0).astype(int)
        # Compute realistic utilization % and workload status matching the mockup
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

    # --- REPORT TAB 6: CUSTOM REPORT BUILDER ---
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
            viz = st.radio(
                "Visualization Type",
                ["Bar Chart", "Column Chart", "Pie Chart", "Line Chart", "Table", "Summary Cards"],
                horizontal=True,
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
                            (report_name.strip(), cfg["type"], json.dumps(cfg), date.today().isoformat(), "James Miller"),
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
# PAGE 8: SETTINGS (Exact Replica of All 8 Mockup Settings Tabs)
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

    # --- SETTINGS TAB 0: WORKSPACE ---
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

    # --- SETTINGS TAB 1: USER MANAGEMENT ---
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

    # --- SETTINGS TAB 2: PROJECT SETTINGS ---
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

    # --- SETTINGS TAB 3: TASK SETTINGS ---
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

    # --- SETTINGS TAB 4: MILESTONES ---
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

    # --- SETTINGS TAB 5: NOTIFICATIONS ---
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

    # --- SETTINGS TAB 6: INTEGRATIONS ---
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

    # --- SETTINGS TAB 7: SYSTEM ---
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
