"""
Vektor — Direction + Velocity. A Microsoft Project-inspired Streamlit project planning app.

Run:
    pip install streamlit pandas plotly openpyxl
    streamlit run project_tracker.py

Data is stored locally in project_tracker.db (SQLite) beside this script.
This is an independent project tracker, not a Microsoft Project replacement or file-format clone.
"""

from __future__ import annotations

import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "project_tracker.db"

TASK_COLUMNS = [
    "task_id", "wbs", "task_name", "owner", "start_date", "end_date",
    "duration_days", "progress", "priority", "status", "predecessors",
    "milestone", "estimated_cost", "actual_cost", "notes"
]
STATUSES = ["Not Started", "In Progress", "On Hold", "Completed", "Cancelled"]
PRIORITIES = ["Low", "Normal", "High", "Critical"]


def db_connect():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db_connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                project_id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT DEFAULT '',
                manager TEXT DEFAULT '',
                start_date TEXT NOT NULL,
                target_date TEXT NOT NULL,
                budget REAL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                task_id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                wbs TEXT DEFAULT '',
                task_name TEXT NOT NULL,
                owner TEXT DEFAULT '',
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                duration_days INTEGER DEFAULT 1,
                progress INTEGER DEFAULT 0,
                priority TEXT DEFAULT 'Normal',
                status TEXT DEFAULT 'Not Started',
                predecessors TEXT DEFAULT '',
                milestone INTEGER DEFAULT 0,
                estimated_cost REAL DEFAULT 0,
                actual_cost REAL DEFAULT 0,
                notes TEXT DEFAULT '',
                FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS resources (
                resource_id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                role TEXT DEFAULT '',
                capacity_pct INTEGER DEFAULT 100,
                hourly_rate REAL DEFAULT 0,
                notes TEXT DEFAULT '',
                FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE,
                UNIQUE(project_id, name)
            )
        """)
        count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
        if count == 0:
            today = date.today()
            cur = conn.execute(
                """INSERT INTO projects(name,description,manager,start_date,target_date,budget,created_at)
                   VALUES(?,?,?,?,?,?,?)""",
                ("Sample Project", "Example project plan — edit or replace with your own.", "",
                 today.isoformat(), (today + timedelta(days=60)).isoformat(), 25000,
                 datetime.now().isoformat(timespec="seconds"))
            )
            pid = cur.lastrowid
            samples = [
                ("1", "Initiation and planning", "", today, today + timedelta(days=4), 100, "Normal", "Completed", "", 0, 1000, 1000, "Confirm scope and stakeholders."),
                ("1.1", "Define project scope", "Project Manager", today, today + timedelta(days=1), 100, "High", "Completed", "", 0, 300, 300, ""),
                ("1.2", "Identify stakeholders", "Project Manager", today + timedelta(days=1), today + timedelta(days=2), 100, "Normal", "Completed", "2", 0, 200, 200, ""),
                ("2", "Design", "Business Analyst", today + timedelta(days=5), today + timedelta(days=12), 30, "High", "In Progress", "2", 0, 4000, 900, "Review design with stakeholders."),
                ("3", "Build", "Development Lead", today + timedelta(days=13), today + timedelta(days=32), 0, "Critical", "Not Started", "4", 0, 12000, 0, ""),
                ("4", "Testing and acceptance", "QA Lead", today + timedelta(days=33), today + timedelta(days=45), 0, "High", "Not Started", "5", 0, 5000, 0, ""),
                ("5", "Go-live", "Project Manager", today + timedelta(days=46), today + timedelta(days=46), 0, "Critical", "Not Started", "6", 1, 1000, 0, "Milestone"),
            ]
            for wbs, name, owner, start, end, progress, priority, status, preds, milestone, est, actual, notes in samples:
                conn.execute("""
                    INSERT INTO tasks(project_id,wbs,task_name,owner,start_date,end_date,duration_days,
                                      progress,priority,status,predecessors,milestone,estimated_cost,actual_cost,notes)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (pid, wbs, name, owner, start.isoformat(), end.isoformat(),
                      max(1, (end-start).days + 1), progress, priority, status, preds, milestone, est, actual, notes))


def get_projects():
    with db_connect() as conn:
        return pd.read_sql_query("SELECT * FROM projects ORDER BY name", conn)


def get_tasks(project_id: int):
    with db_connect() as conn:
        df = pd.read_sql_query("SELECT * FROM tasks WHERE project_id=? ORDER BY task_id", conn, params=(project_id,))
    if df.empty:
        return pd.DataFrame(columns=TASK_COLUMNS)
    for c in ["start_date", "end_date"]:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    return df


def save_project(project_id: int | None, name: str, description: str, manager: str,
                 start: date, target: date, budget: float):
    with db_connect() as conn:
        if project_id is None:
            cur = conn.execute("""INSERT INTO projects(name,description,manager,start_date,target_date,budget,created_at)
                                  VALUES(?,?,?,?,?,?,?)""",
                               (name, description, manager, start.isoformat(), target.isoformat(),
                                budget, datetime.now().isoformat(timespec="seconds")))
            return cur.lastrowid
        conn.execute("""UPDATE projects SET name=?,description=?,manager=?,start_date=?,target_date=?,budget=?
                        WHERE project_id=?""",
                     (name, description, manager, start.isoformat(), target.isoformat(), budget, project_id))
        return project_id


def save_task(project_id: int, row: dict[str, Any], task_id: int | None = None):
    start = row.get("start_date") or date.today()
    end = row.get("end_date") or start
    if isinstance(start, pd.Timestamp):
        start = start.date()
    if isinstance(end, pd.Timestamp):
        end = end.date()
    if end < start:
        raise ValueError("End date cannot be earlier than start date.")
    duration = 1 if bool(row.get("milestone", False)) else max(1, (end - start).days + 1)
    vals = (
        str(row.get("wbs", "") or ""), str(row.get("task_name", "") or "").strip(),
        str(row.get("owner", "") or ""), start.isoformat(), end.isoformat(), duration,
        int(max(0, min(100, float(row.get("progress", 0) or 0)))),
        str(row.get("priority", "Normal") or "Normal"),
        str(row.get("status", "Not Started") or "Not Started"),
        str(row.get("predecessors", "") or ""), int(bool(row.get("milestone", False))),
        float(row.get("estimated_cost", 0) or 0), float(row.get("actual_cost", 0) or 0),
        str(row.get("notes", "") or "")
    )
    if not vals[1]:
        raise ValueError("Task name is required.")
    with db_connect() as conn:
        if task_id:
            conn.execute("""UPDATE tasks SET wbs=?,task_name=?,owner=?,start_date=?,end_date=?,duration_days=?,
                progress=?,priority=?,status=?,predecessors=?,milestone=?,estimated_cost=?,actual_cost=?,notes=?
                WHERE task_id=? AND project_id=?""", vals + (int(task_id), project_id))
        else:
            conn.execute("""INSERT INTO tasks(project_id,wbs,task_name,owner,start_date,end_date,duration_days,
                progress,priority,status,predecessors,milestone,estimated_cost,actual_cost,notes)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (project_id,) + vals)


def delete_task(task_id: int, project_id: int):
    with db_connect() as conn:
        conn.execute("DELETE FROM tasks WHERE task_id=? AND project_id=?", (task_id, project_id))


def get_resources(project_id: int):
    with db_connect() as conn:
        return pd.read_sql_query("SELECT * FROM resources WHERE project_id=? ORDER BY name", conn, params=(project_id,))


def save_resource(project_id: int, name: str, role: str, capacity: int, rate: float, notes: str):
    with db_connect() as conn:
        conn.execute("""INSERT INTO resources(project_id,name,role,capacity_pct,hourly_rate,notes)
                        VALUES(?,?,?,?,?,?)
                        ON CONFLICT(project_id,name) DO UPDATE SET role=excluded.role,
                        capacity_pct=excluded.capacity_pct,hourly_rate=excluded.hourly_rate,notes=excluded.notes""",
                     (project_id, name.strip(), role, capacity, rate, notes))


def delete_resource(resource_id: int, project_id: int):
    with db_connect() as conn:
        conn.execute("DELETE FROM resources WHERE resource_id=? AND project_id=?", (resource_id, project_id))


def calculate_critical_path(tasks: pd.DataFrame):
    """Basic finish-to-start CPM using task IDs in comma-separated predecessor fields."""
    if tasks.empty:
        return set()
    records = {int(r.task_id): r for r in tasks.itertuples(index=False)}
    ids = list(records)
    durations = {i: max(1, int(getattr(records[i], "duration_days", 1) or 1)) for i in ids}
    preds = {}
    for i in ids:
        raw = str(getattr(records[i], "predecessors", "") or "")
        parsed = []
        for item in raw.split(","):
            item = item.strip()
            if item.isdigit() and int(item) in records and int(item) != i:
                parsed.append(int(item))
        preds[i] = parsed
    successors = {i: [] for i in ids}
    for i, ps in preds.items():
        for p in ps:
            successors[p].append(i)

    # Topological ordering; unresolved cycles are excluded from CPM rather than hanging.
    indegree = {i: sum(1 for p in preds[i] if i in successors[p]) for i in ids}
    queue = [i for i in ids if indegree[i] == 0]
    order = []
    while queue:
        node = queue.pop(0)
        order.append(node)
        for nxt in successors[node]:
            indegree[nxt] -= 1
            if indegree[nxt] == 0:
                queue.append(nxt)
    if len(order) != len(ids):
        return set()

    es, ef = {}, {}
    for i in order:
        es[i] = max((ef[p] for p in preds[i]), default=0)
        ef[i] = es[i] + durations[i]
    project_duration = max(ef.values(), default=0)
    lf, ls = {}, {}
    for i in reversed(order):
        lf[i] = min((ls[s] for s in successors[i]), default=project_duration)
        ls[i] = lf[i] - durations[i]
    return {i for i in ids if ls[i] == es[i]}


def fmt_money(value):
    try:
        return f"${float(value):,.2f}"
    except (TypeError, ValueError):
        return "$0.00"


st.set_page_config(page_title="Vektor | Direction + Velocity", page_icon="📊", layout="wide")
init_db()

st.markdown("""
<style>
    :root { --pt-navy:#003B5C; --pt-blue:#01A982; --pt-green:#01A982; --pt-teal:#01A982; --pt-line:#d9e5df; --pt-ink:#173b35; }
    .block-container {padding: 1.1rem 1.35rem 2rem; max-width: 1900px;}
    [data-testid="stMetric"] {background:linear-gradient(145deg,#ffffff,#f3fbf7); border:1px solid var(--pt-line); padding:14px 16px; border-radius:14px; box-shadow:0 2px 8px rgba(0,59,92,.06);}
    [data-testid="stMetricLabel"] {color:#536b63; font-weight:650;}
    [data-testid="stMetricValue"] {color:var(--pt-navy); font-weight:750;}
    .small-muted {color:#687b73; font-size:.85rem;}
    div[data-testid="stDataFrame"], div[data-testid="stTable"] {border-radius:10px; overflow:hidden;}
    .project-title {font-size:clamp(1.45rem,2.2vw,2rem); font-weight:800; letter-spacing:-.03em; color:var(--pt-navy);}
    h1,h2,h3 {color:var(--pt-navy); letter-spacing:-.02em;}
    div[data-testid="stSidebar"] {background:linear-gradient(180deg,#f5fbf8 0%,#eaf5ef 100%); border-right:1px solid #d4e5dc; min-width:255px !important;}
    div[data-testid="stSidebar"] > div {height:100vh; overflow-y:auto;}
    /* Keep the workspace navigation expanded: hide Streamlit's sidebar collapse controls. */
    [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"], button[kind="header"] {display:none !important;}
    div[data-testid="stSidebar"] [data-testid="stRadio"] label {padding:7px 9px; border-radius:8px; font-weight:550;}
    div[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {background:#d8f2e7;}
    div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button {border-radius:9px; font-weight:650; min-height:2.55rem; border-color:#01A982; color:#007F63;}
    div[data-testid="stButton"] > button[kind="primary"] {background:#01A982; color:#ffffff; border-color:#01A982;}
    div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover {background:#e3f7ee; border-color:#01A982; color:#003B5C;}
    [data-testid="stProgressBar"] > div > div {background-color:#01A982;}
    a {color:#01A982;}
    div[data-testid="stPlotlyChart"] {border:1px solid #dce9e2; border-radius:12px; padding:5px; background:#fff;}
    @media (max-width: 900px) {
      .block-container {padding: .7rem .7rem 1.5rem;}
      div[data-testid="stSidebar"] {min-width:210px !important; width:210px !important;}
      [data-testid="stHorizontalBlock"] {flex-wrap:wrap; gap:.55rem !important;}
      [data-testid="stHorizontalBlock"] > [data-testid="column"] {min-width:min(100%, 240px) !important; flex:1 1 240px !important;}
      [data-testid="stMetric"] {padding:10px;}
    }
    @media (max-width: 560px) {
      .block-container {padding-left:.45rem; padding-right:.45rem;}
      h1 {font-size:1.55rem !important;} h2 {font-size:1.25rem !important;} h3 {font-size:1.05rem !important;}
      div[data-testid="stSidebar"] {min-width:190px !important; width:190px !important;}
      button, input, textarea, [role="combobox"] {font-size:16px !important;}
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="project-title">VEKTOR</div>', unsafe_allow_html=True)
st.markdown('<div style="font-size:1rem;font-weight:700;letter-spacing:.12em;color:#01A982;margin-top:-.35rem;margin-bottom:.25rem;">DIRECTION PLUS VELOCITY</div>', unsafe_allow_html=True)
st.caption("Plan • Schedule • Execute • Monitor — a responsive project workspace inspired by Microsoft Project.")

projects = get_projects()
if projects.empty:
    st.error("No projects found. Restart the app to initialize the local database.")
    st.stop()

with st.sidebar:
    st.markdown("## ◈ Project Workspace")
    st.caption("PLAN · TRACK · DELIVER")
    project_options = {f"{r['name']}  ·  #{r['project_id']}": int(r["project_id"]) for _, r in projects.iterrows()}
    selected_label = st.selectbox("ACTIVE PROJECT", list(project_options.keys()))
    project_id = project_options[selected_label]
    project_row = projects.loc[projects["project_id"] == project_id].iloc[0]
    st.markdown("#### WORKSPACE")
    page = st.radio("Navigate", [
        "Overview", "Task Plan & Gantt", "Task Register", "Resources", "Project Settings", "Import / Export"
    ], label_visibility="collapsed")
    st.divider()
    st.caption("● Local database connected")
    st.caption(f"Storage: {DB_PATH.name}")

tasks = get_tasks(project_id)
critical_ids = calculate_critical_path(tasks)
today = date.today()

if page == "Overview":
    st.markdown(f'<div class="project-title">{project_row["name"]}</div>', unsafe_allow_html=True)
    st.write(project_row["description"] or "No project description added.")
    total = len(tasks)
    completed = int((tasks["status"] == "Completed").sum()) if total else 0
    overdue = int(((tasks["end_date"] < today) & (tasks["status"] != "Completed")).sum()) if total else 0
    critical_count = int(tasks["task_id"].isin(critical_ids).sum()) if total else 0
    avg_progress = float(tasks["progress"].mean()) if total else 0
    actual_cost = float(tasks["actual_cost"].sum()) if total else 0
    estimated_cost = float(tasks["estimated_cost"].sum()) if total else 0
    team_count = len(get_resources(project_id))
    open_actions = int((tasks["status"].isin(["Not Started", "In Progress", "On Hold"])).sum()) if total else 0
    high_risk = int((tasks["priority"].isin(["Critical", "High"]) & (tasks["status"] != "Completed")).sum()) if total else 0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("TOTAL TASKS", total)
    c2.metric("TASKS COMPLETED", f"{completed}/{total}")
    c3.metric("OVERDUE TASKS", overdue, delta=f"{overdue} need attention" if overdue else "On schedule", delta_color="inverse")
    c4.metric("PROJECT PROGRESS", f"{avg_progress:.0f}%")
    c5.metric("HIGH-RISK OPEN TASKS", high_risk)
    st.progress(min(1.0, max(0.0, avg_progress / 100)), text=f"Overall task progress: {avg_progress:.1f}%")
    summary_left, summary_mid, summary_right = st.columns(3)
    summary_left.metric("OPEN / PENDING ACTIONS", open_actions)
    summary_mid.metric("TEAM MEMBERS", team_count)
    budget_value = float(project_row["budget"] or 0)
    summary_right.metric("BUDGET UTILIZATION", f"{actual_cost / budget_value:.0%}" if budget_value > 0 else "—")
    left, right = st.columns([1.15, 1])
    with left:
        st.subheader("Schedule")
        if total:
            gantt = tasks.copy()
            gantt["Start"] = pd.to_datetime(gantt["start_date"])
            gantt["Finish"] = pd.to_datetime(gantt["end_date"]) + pd.Timedelta(days=1)
            gantt["Task label"] = gantt.apply(lambda r: f"{r['wbs']}  {r['task_name']}".strip(), axis=1)
            gantt["Health"] = gantt.apply(
                lambda r: "Overdue" if r["end_date"] < today and r["status"] != "Completed"
                else ("Completed" if r["status"] == "Completed" else "In Progress"), axis=1)
            fig = px.timeline(gantt, x_start="Start", x_end="Finish", y="Task label", color="Health",
                              color_discrete_map={"Overdue": "#D64545", "Completed": "#01A982", "In Progress": "#01A982"},
                              hover_data=["owner", "progress", "priority", "status"])
            fig.update_yaxes(autorange="reversed", title="")
            fig.update_layout(height=max(320, min(650, 36 * total + 100)), margin=dict(l=10, r=10, t=20, b=10),
                              legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Add tasks in Task Register to see the schedule.")
    with right:
        st.subheader("Budget snapshot")
        b1, b2 = st.columns(2)
        b1.metric("Estimated cost", fmt_money(estimated_cost))
        b2.metric("Actual cost", fmt_money(actual_cost), delta=fmt_money(actual_cost-estimated_cost), delta_color="inverse")
        budget = float(project_row["budget"] or 0)
        if budget > 0:
            st.progress(min(1.0, actual_cost / budget), text=f"{actual_cost / budget:.0%} of project budget used")
            st.caption(f"Project budget: {fmt_money(budget)}")
        if total:
            st.subheader("Items needing attention")
            attention = tasks[((tasks["end_date"] < today) & (tasks["status"] != "Completed")) |
                              (tasks["priority"].isin(["Critical", "High"]) & (tasks["status"] != "Completed"))]
            if attention.empty:
                st.success("No overdue or high-priority open tasks.")
            else:
                st.dataframe(attention[["task_id", "wbs", "task_name", "owner", "end_date", "priority", "status", "progress"]],
                             hide_index=True, use_container_width=True)
        st.subheader("Status breakdown")
        if total:
            status_df = tasks.groupby("status", dropna=False).size().reset_index(name="Tasks")
            st.plotly_chart(px.pie(status_df, names="status", values="Tasks", hole=.55,
                                   color_discrete_sequence=["#01A982", "#01A982", "#003B5C", "#8ACDB5", "#D64545"]), use_container_width=True)

elif page == "Task Plan & Gantt":
    st.subheader("Task plan and Gantt chart")
    st.caption("Predecessors are comma-separated task IDs and use finish-to-start logic for the basic critical-path estimate.")
    if tasks.empty:
        st.info("Create tasks in Task Register first.")
    else:
        view = st.selectbox("Schedule view", ["All tasks", "Critical path only", "Milestones only", "Open tasks only"])
        display = tasks.copy()
        if view == "Critical path only":
            display = display[display["task_id"].isin(critical_ids)]
        elif view == "Milestones only":
            display = display[display["milestone"] == 1]
        elif view == "Open tasks only":
            display = display[display["status"] != "Completed"]
        if display.empty:
            st.info("No tasks match this view.")
        else:
            plot_df = display.copy()
            plot_df["Start"] = pd.to_datetime(plot_df["start_date"])
            plot_df["Finish"] = pd.to_datetime(plot_df["end_date"]) + pd.Timedelta(days=1)
            plot_df["Task"] = plot_df.apply(lambda r: f"{r['task_id']} · {r['wbs']} {r['task_name']}".strip(), axis=1)
            plot_df["Schedule"] = plot_df["task_id"].apply(lambda x: "Critical path" if int(x) in critical_ids else "Standard")
            fig = px.timeline(plot_df, x_start="Start", x_end="Finish", y="Task", color="Schedule",
                              color_discrete_map={"Critical path": "#01A982", "Standard": "#003B5C"},
                              hover_data=["owner", "progress", "status", "priority", "predecessors"])
            fig.update_yaxes(autorange="reversed", title="")
            fig.update_layout(height=max(360, min(850, 34 * len(plot_df) + 100)), margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)
        st.subheader("Dependency / critical-path summary")
        st.write(f"**{len(critical_ids)}** tasks are on the calculated critical path.")
        st.caption("This is a simplified CPM estimate. It does not yet model resource leveling, calendars, lag/lead, or start-to-start/finish-to-finish links.")

elif page == "Task Register":
    st.subheader("Task register")
    st.caption("Add, edit, or delete tasks. Task IDs are generated automatically; use task IDs in the Predecessors field.")
    with st.expander("＋ Add a task", expanded=False):
        with st.form("add_task_form", clear_on_submit=True):
            a, b, c = st.columns(3)
            wbs = a.text_input("WBS / outline code", placeholder="e.g. 1.2")
            task_name = b.text_input("Task name *")
            owner = c.text_input("Owner / assignee")
            d, e, f = st.columns(3)
            start_date = d.date_input("Start date", value=today)
            end_date = e.date_input("End date", value=today + timedelta(days=1))
            priority = f.selectbox("Priority", PRIORITIES, index=1)
            g, h, i = st.columns(3)
            status = g.selectbox("Status", STATUSES)
            progress = h.slider("Progress (%)", 0, 100, 0)
            milestone = i.checkbox("Milestone (zero-duration marker)")
            j, k, l = st.columns(3)
            predecessors = j.text_input("Predecessors (task IDs)", placeholder="e.g. 2, 4")
            estimated_cost = k.number_input("Estimated cost", min_value=0.0, value=0.0, step=100.0)
            actual_cost = l.number_input("Actual cost", min_value=0.0, value=0.0, step=100.0)
            notes = st.text_area("Notes")
            submitted = st.form_submit_button("Add task", type="primary")
            if submitted:
                try:
                    if end_date < start_date:
                        raise ValueError("End date cannot be earlier than start date.")
                    save_task(project_id, dict(wbs=wbs, task_name=task_name, owner=owner,
                        start_date=start_date, end_date=end_date, priority=priority, status=status,
                        progress=progress, milestone=milestone, predecessors=predecessors,
                        estimated_cost=estimated_cost, actual_cost=actual_cost, notes=notes))
                    st.success("Task added.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
    tasks = get_tasks(project_id)
    if not tasks.empty:
        search = st.text_input("Search tasks", placeholder="Search task, WBS, owner, status…")
        filtered = tasks.copy()
        if search:
            mask = filtered.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
            filtered = filtered[mask]
        visible_cols = ["task_id", "wbs", "task_name", "owner", "start_date", "end_date", "duration_days",
                        "progress", "priority", "status", "predecessors", "milestone", "estimated_cost", "actual_cost", "notes"]
        st.dataframe(filtered[visible_cols], hide_index=True, use_container_width=True,
                     column_config={
                         "task_id": st.column_config.NumberColumn("ID", help="Use this ID for predecessor links."),
                         "progress": st.column_config.ProgressColumn("Progress", min_value=0, max_value=100, format="%d%%"),
                         "milestone": st.column_config.CheckboxColumn("Milestone"),
                         "estimated_cost": st.column_config.NumberColumn("Est. cost", format="$%.2f"),
                         "actual_cost": st.column_config.NumberColumn("Actual cost", format="$%.2f"),
                     })
        st.divider()
        st.subheader("Edit or delete a task")
        task_lookup = {f"#{int(r.task_id)} — {r.task_name}": int(r.task_id) for r in tasks.itertuples(index=False)}
        chosen = st.selectbox("Select task", list(task_lookup.keys()), key="edit_task_select")
        selected_id = task_lookup[chosen]
        row = tasks.loc[tasks["task_id"] == selected_id].iloc[0]
        with st.form(f"edit_task_{selected_id}"):
            a, b, c = st.columns(3)
            ewbs = a.text_input("WBS / outline code", str(row["wbs"] or ""))
            ename = b.text_input("Task name *", str(row["task_name"] or ""))
            eowner = c.text_input("Owner / assignee", str(row["owner"] or ""))
            d, e, f = st.columns(3)
            estart = d.date_input("Start date", value=row["start_date"] or today, key=f"start_{selected_id}")
            eend = e.date_input("End date", value=row["end_date"] or today, key=f"end_{selected_id}")
            epriority = f.selectbox("Priority", PRIORITIES, index=PRIORITIES.index(row["priority"]) if row["priority"] in PRIORITIES else 1)
            g, h, i = st.columns(3)
            estatus = g.selectbox("Status", STATUSES, index=STATUSES.index(row["status"]) if row["status"] in STATUSES else 0)
            eprogress = h.slider("Progress (%)", 0, 100, int(row["progress"]), key=f"progress_{selected_id}")
            emilestone = i.checkbox("Milestone", bool(row["milestone"]), key=f"milestone_{selected_id}")
            j, k, l = st.columns(3)
            epreds = j.text_input("Predecessors (task IDs)", str(row["predecessors"] or ""))
            eest = k.number_input("Estimated cost", min_value=0.0, value=float(row["estimated_cost"] or 0), step=100.0)
            eactual = l.number_input("Actual cost", min_value=0.0, value=float(row["actual_cost"] or 0), step=100.0)
            enotes = st.text_area("Notes", str(row["notes"] or ""))
            saved = st.form_submit_button("Save changes", type="primary")
            if saved:
                try:
                    save_task(project_id, dict(wbs=ewbs, task_name=ename, owner=eowner,
                        start_date=estart, end_date=eend, priority=epriority, status=estatus,
                        progress=eprogress, milestone=emilestone, predecessors=epreds,
                        estimated_cost=eest, actual_cost=eactual, notes=enotes), selected_id)
                    st.success("Task updated.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        confirm_delete = st.checkbox("I understand this permanently deletes the selected task", key=f"confirm_delete_{selected_id}")
        if st.button("Delete selected task", disabled=not confirm_delete):
            delete_task(selected_id, project_id)
            st.success("Task deleted.")
            st.rerun()
    else:
        st.info("No tasks yet. Use Add a task to begin.")

elif page == "Resources":
    st.subheader("Resource management")
    st.caption("Maintain a simple resource directory and capacity/rate reference for this project.")
    resources = get_resources(project_id)
    with st.expander("＋ Add / update resource", expanded=resources.empty):
        with st.form("resource_form", clear_on_submit=True):
            a, b, c = st.columns(3)
            name = a.text_input("Resource name *")
            role = b.text_input("Role")
            capacity = c.number_input("Capacity (%)", min_value=1, max_value=100, value=100)
            d, e = st.columns(2)
            rate = d.number_input("Hourly rate", min_value=0.0, value=0.0, step=5.0)
            notes = e.text_input("Notes")
            if st.form_submit_button("Save resource", type="primary"):
                if not name.strip():
                    st.error("Resource name is required.")
                else:
                    try:
                        save_resource(project_id, name, role, int(capacity), rate, notes)
                        st.success("Resource saved.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
    resources = get_resources(project_id)
    if resources.empty:
        st.info("No resources added yet.")
    else:
        st.dataframe(resources.drop(columns=["project_id"]), hide_index=True, use_container_width=True)
        resource_options = {f"{r['name']} — {r['role']}": int(r["resource_id"]) for _, r in resources.iterrows()}
        selected_resource = st.selectbox("Resource to delete", list(resource_options.keys()))
        if st.button("Delete resource"):
            delete_resource(resource_options[selected_resource], project_id)
            st.rerun()
    if not tasks.empty:
        st.subheader("Assigned workload")
        workload = tasks[tasks["owner"].astype(str).str.strip() != ""].groupby("owner").agg(
            Tasks=("task_id", "count"), AverageProgress=("progress", "mean"),
            EstimatedCost=("estimated_cost", "sum"), ActualCost=("actual_cost", "sum")
        ).reset_index()
        if not workload.empty:
            workload["AverageProgress"] = workload["AverageProgress"].round(1)
            st.dataframe(workload, hide_index=True, use_container_width=True)

elif page == "Project Settings":
    st.subheader("Project settings")
    with st.form("project_settings"):
        name = st.text_input("Project name", str(project_row["name"]))
        manager = st.text_input("Project manager", str(project_row["manager"] or ""))
        description = st.text_area("Description", str(project_row["description"] or ""))
        a, b, c = st.columns(3)
        pstart = a.date_input("Project start", value=date.fromisoformat(project_row["start_date"]))
        ptarget = b.date_input("Target finish", value=date.fromisoformat(project_row["target_date"]))
        budget = c.number_input("Budget", min_value=0.0, value=float(project_row["budget"] or 0), step=1000.0)
        if st.form_submit_button("Save project settings", type="primary"):
            if not name.strip():
                st.error("Project name is required.")
            elif ptarget < pstart:
                st.error("Target finish cannot be earlier than project start.")
            else:
                try:
                    save_project(project_id, name.strip(), description, manager, pstart, ptarget, budget)
                    st.success("Project settings saved.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("A project with that name already exists.")
    st.divider()
    st.subheader("Create another project")
    with st.form("new_project"):
        new_name = st.text_input("New project name")
        new_desc = st.text_area("Description", key="new_desc")
        new_manager = st.text_input("Project manager", key="new_manager")
        a, b = st.columns(2)
        new_start = a.date_input("Start date", value=today, key="new_start")
        new_target = b.date_input("Target finish", value=today + timedelta(days=30), key="new_target")
        new_budget = st.number_input("Budget", min_value=0.0, value=0.0, step=1000.0, key="new_budget")
        if st.form_submit_button("Create project"):
            if not new_name.strip():
                st.error("Project name is required.")
            elif new_target < new_start:
                st.error("Target finish cannot be earlier than project start.")
            else:
                try:
                    save_project(None, new_name.strip(), new_desc, new_manager, new_start, new_target, new_budget)
                    st.success("Project created. Select it from the sidebar.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("A project with that name already exists.")

elif page == "Import / Export":
    st.subheader("Import and export")
    st.write("Export the active project for reporting or backup. CSV is useful for bulk editing; Excel includes separate project, task, and resource sheets.")
    resources = get_resources(project_id)
    export_tasks = get_tasks(project_id)
    export_project = projects[projects["project_id"] == project_id]
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Download tasks as CSV**")
        st.download_button("Download task register (.csv)", export_tasks.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"project_{project_id}_tasks.csv", mime="text/csv")
        st.markdown("**Download project workbook**")
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export_project.to_excel(writer, index=False, sheet_name="Project")
            export_tasks.to_excel(writer, index=False, sheet_name="Tasks")
            resources.to_excel(writer, index=False, sheet_name="Resources")
        st.download_button("Download Excel workbook (.xlsx)", buffer.getvalue(),
                           file_name=f"project_{project_id}_tracker.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.markdown("**Import tasks from CSV**")
        st.caption("Expected columns include task_name, start_date, end_date. Optional columns are matched by name.")
        uploaded = st.file_uploader("Choose task CSV", type=["csv"], key="task_csv_upload")
        if uploaded is not None:
            try:
                incoming = pd.read_csv(uploaded)
                st.write("Preview", incoming.head(10))
                if st.button("Import previewed tasks", type="primary"):
                    if "task_name" not in incoming.columns:
                        st.error("CSV must contain a task_name column.")
                    else:
                        added, errors = 0, []
                        for idx, r in incoming.iterrows():
                            try:
                                sd = pd.to_datetime(r.get("start_date", today), errors="coerce")
                                ed = pd.to_datetime(r.get("end_date", sd), errors="coerce")
                                if pd.isna(sd):
                                    sd = pd.Timestamp(today)
                                if pd.isna(ed):
                                    ed = sd
                                data = {}
                                for col in ["wbs", "task_name", "owner", "priority", "status", "predecessors", "milestone", "estimated_cost", "actual_cost", "notes", "progress"]:
                                    val = r.get(col, None)
                                    if pd.isna(val):
                                        val = None
                                    data[col] = val
                                data["start_date"], data["end_date"] = sd.date(), ed.date()
                                data["milestone"] = bool(data.get("milestone", False)) if data.get("milestone") is not None else False
                                save_task(project_id, data)
                                added += 1
                            except Exception as exc:
                                errors.append(f"Row {idx + 2}: {exc}")
                        st.success(f"Imported {added} task(s).")
                        if errors:
                            st.warning("Some rows were skipped:\n\n" + "\n".join(errors[:10]))
                        st.rerun()
            except Exception as exc:
                st.error(f"Could not read CSV: {exc}")
    st.divider()
    st.markdown("**Database backup**")
    if DB_PATH.exists():
        st.download_button("Download full SQLite database", DB_PATH.read_bytes(),
                           file_name="project_tracker_backup.db", mime="application/octet-stream")
    st.caption("Keep a secure backup of the SQLite database. This local app does not include multi-user authentication or cloud synchronization.")
