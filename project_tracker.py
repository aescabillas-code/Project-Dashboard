"""Vektor Project Management - High-Fidelity Streamlit Application.
Install: pip install streamlit pandas plotly openpyxl pymongo dnspython
Run: streamlit run project_tracker.py
"""
from __future__ import annotations

import base64
import calendar
import html
import io
import json
import os
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pymongo import MongoClient

DB_NAME = "Vektor"
COLLECTION_NAME = "Reports"

@st.cache_resource
def get_mongo_collection():
    try:
        # Securely retrieve the connection string from Streamlit secrets
        mongo_uri = st.secrets["MONGO_URI"]
        
        # Initialize MongoClient with connection timeout
        client = MongoClient(
            mongo_uri, 
            serverSelectionTimeoutMS=5000
        )
        db = client[DB_NAME]
        return db[COLLECTION_NAME]
    except Exception as e:
        st.error(f"MongoDB Connection Error: {e}")
        return None

collection = get_mongo_collection()

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
# Data Layer with MongoDB Persistence & Automated Calculations
# ---------------------------------------------------------------------------
def load_data_from_mongo():
    if collection is not None:
        try:
            doc = collection.find_one({"_id": "vektor_state"})
            if doc and "projects" in doc:
                return doc["projects"], doc["tasks"], doc["resources"], doc["milestones"], doc["settings"], doc.get("saved_reports", [])
        except Exception:
            pass
    return None, None, None, None, None, None

def save_data_to_mongo(projects, tasks, resources, milestones, settings, saved_reports):
    if collection is not None:
        try:
            payload = {
                "_id": "vektor_state",
                "projects": projects,
                "tasks": tasks,
                "resources": resources,
                "milestones": milestones,
                "settings": settings,
                "saved_reports": saved_reports,
                "updated_at": datetime.utcnow().isoformat()
            }
            collection.replace_one({"_id": "vektor_state"}, payload, upsert=True)
        except Exception as e:
            st.error(f"Failed to save data to MongoDB: {e}")

# Initial Seeding / Temporary Mock Data (Disappears once real DB data is entered/saved)
default_projects = [
    {
        "id": 1, "name": "Website Revamp", "owner": "Alice Morgan", "team": "Marketing", "category": "Web Development",
        "status": "On Track", "priority": "High", "start_date": "2024-04-24", "end_date": "2024-06-30",
        "progress": 70, "days_left": 45, "tasks_done": 14, "tasks_total": 20, "tags": "Web,Marketing,Corporate",
        "bar_color": "#00875A", "short_desc": "Redesign corporate website", "description": "Redesign and launch the corporate website with improved user experience.", "health": "Healthy"
    },
    {
        "id": 2, "name": "E-Commerce Platform", "owner": "Brian Taylor", "team": "IT", "category": "IT & Development",
        "status": "In Progress", "priority": "High", "start_date": "2024-03-15", "end_date": "2024-06-15",
        "progress": 50, "days_left": 30, "tasks_done": 12, "tasks_total": 24, "tags": "IT,Development",
        "bar_color": "#1473E6", "short_desc": "Build new e-commerce platform", "description": "Build new e-commerce platform to expand digital presence.", "health": "Healthy"
    },
    {
        "id": 3, "name": "Mobile App Development", "owner": "Emily Rodriguez", "team": "Design", "category": "Mobile Development",
        "status": "At Risk", "priority": "Medium", "start_date": "2024-04-01", "end_date": "2024-08-30",
        "progress": 30, "days_left": 90, "tasks_done": 8, "tasks_total": 26, "tags": "Mobile,Customer",
        "bar_color": "#7C3AED", "short_desc": "Develop customer mobile app", "description": "Develop customer mobile application for iOS and Android.", "health": "At Risk"
    },
    {
        "id": 4, "name": "Data Migration", "owner": "Carlos Diaz", "team": "IT", "category": "Cloud & Data",
        "status": "On Track", "priority": "High", "start_date": "2024-05-01", "end_date": "2024-07-31",
        "progress": 80, "days_left": 60, "tasks_done": 18, "tasks_total": 22, "tags": "Data,Cloud",
        "bar_color": "#F28C00", "short_desc": "Migrate legacy data to cloud", "description": "Migrate legacy data to cloud infrastructure with minimal downtime.", "health": "Healthy"
    },
    {
        "id": 5, "name": "IT Infrastructure Upgrade", "owner": "Dana Scott", "team": "IT", "category": "Infrastructure",
        "status": "At Risk", "priority": "High", "start_date": "2024-05-20", "end_date": "2024-09-15",
        "progress": 20, "days_left": 110, "tasks_done": 6, "tasks_total": 28, "tags": "Infrastructure,IT",
        "bar_color": "#00875A", "short_desc": "Upgrade servers and network", "description": "Upgrade servers and network infrastructure to support business growth.", "health": "At Risk"
    }
]

default_tasks = [
    {"id": 1, "name": "UI/UX Design Complete", "project_id": 1, "team": "Design", "status": "Completed", "priority": "High", "start_date": "2024-04-01", "due_date": "2024-04-30", "assignee": "Alice Morgan", "progress": 100, "dependency": "", "notes": "Design phase completed."},
    {"id": 2, "name": "Homepage Development", "project_id": 1, "team": "IT", "status": "In Progress", "priority": "High", "start_date": "2024-04-10", "due_date": "2024-05-15", "assignee": "James Miller", "progress": 65, "dependency": "UI/UX Design Complete", "notes": "Frontend responsive build."},
    {"id": 3, "name": "Content Review", "project_id": 7, "team": "Marketing", "status": "In Progress", "priority": "Medium", "start_date": "2024-04-15", "due_date": "2024-05-20", "assignee": "Sarah Kim", "progress": 55, "dependency": "", "notes": "Copywriting review."},
    {"id": 4, "name": "API Integration", "project_id": 2, "team": "IT", "status": "In Progress", "priority": "High", "start_date": "2024-04-12", "due_date": "2024-05-25", "assignee": "Brian Taylor", "progress": 60, "dependency": "", "notes": "Payment gateway & cart API."}
]

default_resources = [
    {"id": 1, "name": "Alice Morgan", "email": "alice.morgan@vektor.example", "team": "Marketing", "role": "Admin", "status": "Available", "capacity": 100, "notes": ""},
    {"id": 2, "name": "Brian Taylor", "email": "brian.taylor@vektor.example", "team": "IT", "role": "Project Manager", "status": "Available", "capacity": 100, "notes": ""},
    {"id": 3, "name": "Emily Rodriguez", "email": "emily.rodriguez@vektor.example", "team": "Design", "role": "Project Manager", "status": "Partial", "capacity": 50, "notes": ""},
    {"id": 4, "name": "James Miller", "email": "james.miller@vektor.example", "team": "IT", "role": "Team Member", "status": "Available", "capacity": 100, "notes": ""}
]

default_milestones = [
    {"id": 1, "name": "Project Kick-off", "project_id": 1, "team": "Marketing", "status": "Completed", "start_date": "2024-04-20", "due_date": "2024-04-24", "completion_date": "2024-04-24", "dependency": "", "notes": ""},
    {"id": 2, "name": "Design Sign-off", "project_id": 1, "team": "Design", "status": "Completed", "start_date": "2024-04-10", "due_date": "2024-04-30", "completion_date": "2024-04-30", "dependency": "Project Kick-off", "notes": ""}
]

default_settings = {
    "workspace_name": "Vektor",
    "workspace_description": "Project management workspace for planning, tracking, and delivery.",
    "theme": "HPE Green",
    "language": "English",
    "date_format": "MMM DD, YYYY",
    "default_landing": "Overview",
    "audit_logs": "true"
}

# Load from Mongo or initialize default state
m_proj, m_task, m_res, m_ms, m_set, m_rep = load_data_from_mongo()
if m_proj is None:
    m_proj = default_projects
    m_task = default_tasks
    m_res = default_resources
    m_ms = default_milestones
    m_set = default_settings
    m_rep = []
    save_data_to_mongo(m_proj, m_task, m_res, m_ms, m_set, m_rep)

# Automated Calculations / Real-time Status Synchronization
def run_automations():
    global m_proj, m_task
    for p in m_proj:
        p_tasks = [t for t in m_task if t.get("project_id") == p["id"]]
        if p_tasks:
            tot = len(p_tasks)
            done = len([t for t in p_tasks if t.get("status") == "Completed"])
            avg_prog = int(sum([t.get("progress", 0) for t in p_tasks]) / tot)
            p["tasks_total"] = tot
            p["tasks_done"] = done
            p["progress"] = avg_prog
            if avg_prog == 100:
                p["status"] = "Completed"
                p["health"] = "Healthy"
            elif avg_prog > 0 and p["status"] == "Not Started":
                p["status"] = "In Progress"
    save_data_to_mongo(m_proj, m_task, m_res, m_ms, m_set, m_rep)

run_automations()

P = pd.DataFrame(m_proj)
T = pd.DataFrame(m_task)
R = pd.DataFrame(m_res)
M = pd.DataFrame(m_ms)

# Map project names
if not T.empty and not P.empty and "project_id" in T.columns:
    pid_map = dict(zip(P.id, P.name))
    T["project"] = T["project_id"].map(pid_map)
if not M.empty and not P.empty and "project_id" in M.columns:
    pid_map = dict(zip(P.id, P.name))
    M["project"] = M["project_id"].map(pid_map)

projects_list = P.name.tolist() if not P.empty else []
pmap = dict(zip(P.name, P.id)) if not P.empty else {}
teams_list = sorted(list(set(R.team.dropna().tolist() + P.team.dropna().tolist()))) if not R.empty and not P.empty else ["IT", "Marketing", "Design", "Operations"]

# Theme Configuration
active_theme = m_set.get("theme", "HPE Green") if isinstance(m_set, dict) else "HPE Green"
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
# Global CSS Styling (Inter font, compact tables, comfortable font sizing)
# ---------------------------------------------------------------------------
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

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
    font-size: 14.5px !important;
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
    font-size: 13.5px !important;
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

/* Compact Table Row Spacing */
.vk-table-wrap {{
    overflow-x: auto;
    border-radius: 8px;
    border: 1px solid #EDF2F7;
    background: {CARD_BG};
}}
table.vk-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    text-align: left;
}}
table.vk-table thead tr {{
    background: #F8FAFC;
    border-bottom: 1px solid #E2E8F0;
}}
table.vk-table th {{
    padding: 6px 10px !important;
    font-weight: 600;
    font-size: 12px;
    color: #475569;
}}
table.vk-table td {{
    padding: 5px 10px !important;
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
    padding: 2px 8px;
    border-radius: 6px;
    font-size: 11px;
    font-weight: 600;
    white-space: nowrap;
}}
.vk-badge-green {{ background: #DCFCE7; color: #15803D; }}
.vk-badge-blue {{ background: #DBEAFE; color: #1D4ED8; }}
.vk-badge-orange {{ background: #FEF3C7; color: #B45309; }}
.vk-badge-red {{ background: #FEE2E2; color: #DC2626; }}
.vk-badge-gray {{ background: #E2E8F0; color: #475569; }}
.vk-badge-purple {{ background: #F3E8FF; color: #7C3AED; }}

.vk-tag {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 5px;
    font-size: 11px;
    font-weight: 600;
    margin-right: 4px;
}}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Sidebar Navigation
# ---------------------------------------------------------------------------
with st.sidebar:
    ws_title = m_set.get("workspace_name", "Vektor") if isinstance(m_set, dict) else "Vektor"
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
    default_page = m_set.get("default_landing", "Overview") if isinstance(m_set, dict) else "Overview"
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


# ---------------------------------------------------------------------------
# UI Helpers & Badges
# ---------------------------------------------------------------------------
def badge_html(val: object) -> str:
    s = str(val or "Unknown")
    mapping = {
        "On Track": "vk-badge-green", "Completed": "vk-badge-green", "Healthy": "vk-badge-green", "Available": "vk-badge-green",
        "In Progress": "vk-badge-blue", "At Risk": "vk-badge-orange", "High": "vk-badge-red", "Delayed": "vk-badge-red",
        "Overdue": "vk-badge-red", "On Leave": "vk-badge-red", "Not Started": "vk-badge-gray"
    }
    return f'<span class="vk-badge {mapping.get(s, "vk-badge-gray")}">{html.escape(s)}</span>'


# ===========================================================================
# MAIN ROUTER: RENDER ACTIVE PAGE
# ===========================================================================
if page == "Overview":
    st.markdown("## Project Overview[cite: 35]")
    st.write("Track progress, monitor tasks, and keep your projects on schedule[cite: 35].")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        metric_card("Total Projects", len(P), "Active Portfolio", "📁", "green")
    with col2:
        metric_card("Total Tasks", len(T), "Recorded Items", "☑️", "blue")
    with col3:
        metric_card("Team Resources", len(R), "Available Members", "👥", "purple")
    with col4:
        avg_p = int(P.progress.mean()) if not P.empty else 0
        metric_card("Average Progress", f"{avg_p}%", "Overall completion", "📈", "orange")

    st.write("")
    with st.container(border=True):
        st.markdown("### Active Projects Summary[cite: 35]")
        if P.empty:
            st.info("No projects recorded yet. Add your first project in the Projects tab.")
        else:
            render_rich_table(P, [("name", "Project Name", "bold"), ("owner", "Owner", "avatar"), ("team", "Team", "text"), ("status", "Status", "badge"), ("progress", "Progress", "progress"), ("days_left", "Days Left", "bold")])

elif page == "Projects":
    st.markdown("## Projects Management[cite: 35]")
    c1, c2 = st.columns([4, 1])
    with c2:
        if st.button("＋ Add Project", type="primary", use_container_width=True):
            new_p = {"id": len(m_proj)+1, "name": f"New Project {len(m_proj)+1}", "owner": "Alice Morgan", "team": "IT", "category": "Development", "status": "On Track", "priority": "High", "start_date": str(date.today()), "end_date": str(date.today() + timedelta(days=30)), "progress": 0, "days_left": 30, "tasks_done": 0, "tasks_total": 10, "tags": "Web", "bar_color": GREEN, "short_desc": "New project description", "description": "Detailed project description.", "health": "Healthy"}
            m_proj.append(new_p)
            save_data_to_mongo(m_proj, m_task, m_res, m_ms, m_set, m_rep)
            st.rerun()

    with st.container(border=True):
        render_rich_table(P, [("id", "ID", "text"), ("name", "Project Name", "bold"), ("owner", "Manager", "avatar"), ("category", "Category", "text"), ("status", "Status", "badge"), ("priority", "Priority", "badge"), ("progress", "Progress", "progress")])

elif page == "Gantt Chart":
    st.markdown("## Gantt Chart & Schedule[cite: 35]")
    with st.container(border=True):
        if P.empty:
            st.info("No schedule data available.")
        else:
            fig = px.timeline(P, x_start="start_date", x_end="end_date", y="name", color="status", color_discrete_map={"On Track": GREEN, "At Risk": ORANGE})
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)

elif page == "Task List":
    st.markdown("## Task List Management[cite: 35]")
    c1, c2 = st.columns([4, 1])
    with c2:
        if st.button("＋ Add Task", type="primary", use_container_width=True):
            new_t = {"id": len(m_task)+1, "name": f"New Task {len(m_task)+1}", "project_id": 1, "team": "IT", "status": "Not Started", "priority": "Medium", "start_date": str(date.today()), "due_date": str(date.today() + timedelta(days=7)), "assignee": "James Miller", "progress": 0, "dependency": "", "notes": "Task notes."}
            m_task.append(new_t)
            run_automations()
            st.rerun()

    with st.container(border=True):
        if T.empty:
            st.info("No tasks available.")
        else:
            render_rich_table(T, [("name", "Task Name", "bold"), ("project", "Project", "text"), ("status", "Status", "badge"), ("priority", "Priority", "badge"), ("assignee", "Assignee", "avatar"), ("progress", "Progress", "progress"), ("due_date", "Due Date", "date")])

elif page == "Resources":
    st.markdown("## Resource Directory & Allocation[cite: 35]")
    c1, c2 = st.columns([4, 1])
    with c2:
        if st.button("＋ Add Resource", type="primary", use_container_width=True):
            new_r = {"id": len(m_res)+1, "name": f"New Member {len(m_res)+1}", "email": "member@vektor.example", "team": "IT", "role": "Team Member", "status": "Available", "capacity": 100, "notes": ""}
            m_res.append(new_r)
            save_data_to_mongo(m_proj, m_task, m_res, m_ms, m_set, m_rep)
            st.rerun()

    with st.container(border=True):
        if R.empty:
            st.info("No resources found.")
        else:
            render_rich_table(R, [("name", "Name", "avatar"), ("email", "Email", "text"), ("team", "Team", "text"), ("role", "Role", "badge"), ("status", "Availability", "badge"), ("capacity", "Capacity (%)", "bold")])

elif page == "Calendar":
    st.markdown("## Calendar Schedule[cite: 35]")
    with st.container(border=True):
        st.info("Interactive calendar view synchronized with project milestones and task due dates.")
        if not M.empty:
            render_rich_table(M, [("name", "Milestone Name", "bold"), ("project", "Project", "text"), ("due_date", "Due Date", "date"), ("status", "Status", "badge")])

elif page == "Reports":
    st.markdown("## Automated Reports & Data Analytics[cite: 35]")
    col_exp, _ = st.columns([1, 3])
    with col_exp:
        st.download_button(
            "⬇ Export Report (.xlsx)[cite: 35]",
            excel_bytes({"Projects": P, "Tasks": T, "Resources": R, "Milestones": M}),
            file_name="vektor_automated_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    with st.container(border=True):
        st.markdown("### Portfolio Health & Progress Analytics[cite: 35]")
        c1, c2 = st.columns(2)
        with c1:
            chart_status_donut(P, "status", "Projects")
        with c2:
            chart_status_donut(T, "status", "Tasks")

elif page == "Settings":
    st.markdown("## Workspace Settings & Configuration[cite: 35]")
    with st.container(border=True):
        ws_name = st.text_input("Workspace Name", value=m_set.get("workspace_name", "Vektor"))
        ws_desc = st.text_area("Workspace Description", value=m_set.get("workspace_description", ""))
        theme_sel = st.selectbox("Theme Appearance", ["HPE Green", "Blue", "Dark"], index=0 if m_set.get("theme")=="HPE Green" else 1)
        
        if st.button("💾 Save Settings to MongoDB", type="primary"):
            m_set["workspace_name"] = ws_name
            m_set["workspace_description"] = ws_desc
            m_set["theme"] = theme_sel
            save_data_to_mongo(m_proj, m_task, m_res, m_ms, m_set, m_rep)
            st.success("Changes and configuration saved successfully to MongoDB Atlas collection[cite: 35].")
            st.rerun()
