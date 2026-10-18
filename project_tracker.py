"""Vektor Project Management - Streamlit application.
Install: pip install streamlit pandas plotly openpyxl
Run: streamlit run vektor_project_management_app.py
Uses a local SQLite database (vektor.db), seeded with demo data on first run.
"""
from __future__ import annotations
import sqlite3, io, json, os, calendar
from datetime import date, datetime, timedelta
import pandas as pd
import plotly.express as px
import streamlit as st

DB = os.environ.get('VEKTOR_DB_PATH', 'vektor.db')
GREEN, BLUE, ORANGE, RED, GRAY = '#00875A', '#1473E6', '#F28C00', '#EF4444', '#8495B0'
NAVY = '#10194F'
st.set_page_config(page_title='Vektor | Project Management', page_icon='📊', layout='wide')

def conn():
    c=sqlite3.connect(DB, check_same_thread=False); c.row_factory=sqlite3.Row; return c
def sql(q,p=()):
    with conn() as c: c.execute(q,p); c.commit()
def df(q,p=()):
    with conn() as c: return pd.read_sql_query(q,c,params=p)
def init():
    with conn() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY,name TEXT NOT NULL,owner TEXT,team TEXT,status TEXT,start_date TEXT,end_date TEXT,progress INTEGER,description TEXT,health TEXT);
        CREATE TABLE IF NOT EXISTS tasks(id INTEGER PRIMARY KEY,name TEXT NOT NULL,project_id INTEGER,team TEXT,status TEXT,priority TEXT,start_date TEXT,due_date TEXT,assignee TEXT,progress INTEGER,dependency TEXT,notes TEXT);
        CREATE TABLE IF NOT EXISTS resources(id INTEGER PRIMARY KEY,name TEXT NOT NULL,email TEXT,team TEXT,role TEXT,status TEXT,capacity INTEGER,notes TEXT);
        CREATE TABLE IF NOT EXISTS milestones(id INTEGER PRIMARY KEY,name TEXT NOT NULL,project_id INTEGER,team TEXT,status TEXT,start_date TEXT,due_date TEXT,completion_date TEXT,dependency TEXT,notes TEXT);
        CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
        CREATE TABLE IF NOT EXISTS saved_reports(id INTEGER PRIMARY KEY,name TEXT,report_type TEXT,config TEXT,last_generated TEXT,created_by TEXT);
        CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY,event TEXT,item_type TEXT,item_id INTEGER,details TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        ''')
        if c.execute('SELECT COUNT(*) FROM projects').fetchone()[0]==0:
            ps=[('Website Revamp','James Miller','IT','On Track','2024-04-01','2024-09-30',90,'Website redesign','Healthy'),('Marketing Campaign','Sarah Kim','Marketing','At Risk','2024-04-15','2024-08-31',60,'Integrated campaign','At Risk'),('IT Infrastructure Upgrade','Michael Tan','IT','At Risk','2024-05-01','2024-10-31',40,'Infrastructure modernization','At Risk'),('Sustainability Initiative','Alice Morgan','Operations','Delayed','2024-05-15','2024-12-15',25,'Sustainability program','Critical'),('Internal Operations','Ryan Cruz','Operations','On Track','2024-06-01','2024-11-30',75,'Process improvements','Healthy')]
            c.executemany('INSERT INTO projects(name,owner,team,status,start_date,end_date,progress,description,health) VALUES(?,?,?,?,?,?,?,?,?)',ps)
            people=['James Miller','Sarah Kim','Alice Morgan','Ryan Cruz','Michael Tan','Daniel Perez','Ella Ramirez','Chris Padilla','Anna Reyes','Kevin Santos','Lisa Chan','David Lim','Maria Santos','John Cruz','Paolo Reyes','Nina Flores','Mark Lee','Jade Tan','Ben Garcia','Mia Lopez','Leo Ramos','Ivy Chen','Noah King','Zoe Park']
            teams=['IT','Design','Marketing','Operations','IT','Operations','Design','IT','Marketing','Operations','Design','IT','Marketing','Operations','IT','Design','Marketing','Operations','IT','Design','IT','Marketing','Operations','Others']
            for i,(n,t) in enumerate(zip(people,teams)):
                status='On Leave' if i in (4,10) else ('Partial' if i in (2,6,14,17) else 'Available')
                c.execute('INSERT INTO resources(name,email,team,role,status,capacity) VALUES(?,?,?,?,?,?)',(n,n.lower().replace(' ','.')+'@vektor.example',t,'Admin' if i==0 else ('Project Manager' if i in (1,4) else 'Team Member'),status,0 if status=='On Leave' else (50 if status=='Partial' else 100)))
            tasks=[('UI/UX Design Complete',1,'Design','Completed','High','2024-04-01','2024-04-30','Alice Morgan',100,'','Design phase completed.'),('Homepage Development',1,'IT','In Progress','High','2024-04-10','2024-05-15','James Miller',65,'UI/UX Design Complete',''),('Content Review',2,'Marketing','In Progress','Medium','2024-04-15','2024-05-20','Sarah Kim',55,'',''),('Infrastructure Setup',3,'IT','Not Started','High','2024-05-01','2024-05-15','Michael Tan',0,'Requirements Gathering',''),('Sustainability Plan',4,'Operations','Completed','Medium','2024-05-05','2024-05-30','Ryan Cruz',100,'',''),('Process Documentation',5,'Operations','In Progress','Low','2024-05-10','2024-06-20','Daniel Perez',50,'',''),('Training Materials',2,'Marketing','Not Started','Medium','2024-05-20','2024-06-30','Ella Ramirez',0,'',''),('QA Testing',1,'IT','In Progress','High','2024-05-25','2024-06-25','Chris Padilla',35,'Homepage Development',''),('Campaign Launch',2,'Marketing','Not Started','High','2024-05-25','2024-06-10','Sarah Kim',0,'Content Review',''),('System Configuration',3,'IT','Completed','Low','2024-03-15','2024-04-10','Michael Tan',100,'',''),('Requirements Gathering',5,'Operations','Completed','Low','2024-04-01','2024-04-25','Daniel Perez',100,'','')]
            c.executemany('INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)',tasks)
            ms=[('UI/UX Design Complete',1,'Design','Completed','2024-04-01','2024-04-30','2024-04-28','','Design phase completed.'),('Campaign Launch',2,'Marketing','In Progress','2024-04-15','2024-05-20','','Content Review','Awaiting final assets.'),('Infrastructure Setup',3,'IT','In Progress','2024-05-01','2024-05-15','','Requirements Gathering','Environment setup ongoing.'),('Sustainability Plan',4,'Operations','Not Started','2024-05-05','2024-05-25','','','Pending approval.'),('Process Documentation',5,'Operations','Not Started','2024-05-10','2024-05-30','','Infrastructure Setup',''),('QA Testing',1,'IT','Overdue','2024-04-01','2024-04-20','','Development Complete','Testing issues to resolve.'),('Training Materials',2,'Marketing','Overdue','2024-04-20','2024-05-05','','Content Review',''),('System Configuration',3,'IT','Completed','2024-03-15','2024-04-10','2024-04-09','','Configuration completed.')]
            c.executemany('INSERT INTO milestones(name,project_id,team,status,start_date,due_date,completion_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?)',ms)
            c.executemany('INSERT INTO settings(key,value) VALUES(?,?)',[('workspace_name','Vektor'),('theme','HPE Green'),('default_landing','Overview')]); c.commit()
init()
P=df('SELECT * FROM projects ORDER BY id')
T=df('SELECT t.*,p.name project FROM tasks t LEFT JOIN projects p ON p.id=t.project_id ORDER BY t.due_date')
R=df('SELECT * FROM resources ORDER BY name')
M=df('SELECT m.*,p.name project FROM milestones m LEFT JOIN projects p ON p.id=m.project_id ORDER BY m.due_date')
projects=P.name.tolist(); pmap=dict(zip(P.name,P.id)); teams=sorted(set(R.team.dropna().tolist()+P.team.dropna().tolist()))

st.markdown('''<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:Inter,Arial,sans-serif}.stApp{background:linear-gradient(115deg,#F3FAFF,#EAF5FC);color:#10194F}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#001F1D,#002F2A)}[data-testid="stSidebar"] *{color:#F4FBFA!important}
.block-container{padding-top:1.2rem;max-width:1800px}h1,h2,h3{color:#10194F!important;letter-spacing:-.035em}h1{font-weight:800!important}
[data-testid="stMetric"]{background:white;border:1px solid #e1edf6;border-radius:12px;padding:14px;box-shadow:0 2px 10px #12325a08}
.stButton>button,.stDownloadButton>button{border-radius:8px;font-weight:600;border:1px solid #c8d9e7}.stButton>button[kind="primary"]{background:#00875A;color:white;border-color:#00875A}
div[data-testid="stVerticalBlockBorderWrapper"]{background:#ffffffed;border:1px solid #e1edf6;border-radius:12px;padding:12px}
.metric-card{display:flex;align-items:center;gap:12px;background:white;border:1px solid #e1edf6;border-radius:12px;padding:14px;min-height:92px}.metric-icon{width:48px;height:48px;display:flex;align-items:center;justify-content:center;border-radius:50%;background:#d9f3ef;font-size:24px}.metric-label{font-size:12px;font-weight:650}.metric-value{font-size:27px;font-weight:800;line-height:1.2}.metric-note{font-size:11px;color:#61718b;margin-top:3px}.section-title{font-size:16px;font-weight:750;color:#10194F;margin-bottom:8px}
</style>''',unsafe_allow_html=True)

def log(event,typ='',item_id=None,details=''): sql('INSERT INTO audit_log(event,item_type,item_id,details) VALUES(?,?,?,?)',(event,typ,item_id,details))
def metric(label,value,note='',icon='📊'):
    st.markdown(f'<div class="metric-card"><div class="metric-icon">{icon}</div><div><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div class="metric-note">{note}</div></div></div>',unsafe_allow_html=True)
def section(title): st.markdown(f'<div class="section-title">{title}</div>',unsafe_allow_html=True)
def chart_status(data,col):
    if data.empty or col not in data: st.info('No data available.'); return
    z=data[col].fillna('Unknown').value_counts().rename_axis('Status').reset_index(name='Count')
    colors={'On Track':GREEN,'Healthy':GREEN,'Completed':GREEN,'In Progress':BLUE,'At Risk':ORANGE,'Partial':ORANGE,'Delayed':RED,'Critical':RED,'Not Started':GRAY,'Available':GREEN,'On Leave':RED,'Overdue':RED}
    fig=px.pie(z,names='Status',values='Count',hole=.58,color='Status',color_discrete_map=colors)
    fig.update_layout(height=250,margin=dict(l=0,r=0,t=0,b=0),paper_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
def progress_chart(data,name='name',val='progress'):
    if data.empty: st.info('No data available.'); return
    z=data[[name,val]].copy(); z[val]=pd.to_numeric(z[val],errors='coerce').fillna(0)
    fig=px.bar(z,x=val,y=name,orientation='h',range_x=[0,100],text=val,color=val,color_continuous_scale=['#EF4444','#F28C00','#00875A'])
    fig.update_traces(texttemplate='%{text}%',textposition='outside',cliponaxis=False)
    fig.update_layout(height=max(220,len(z)*36),margin=dict(l=0,r=24,t=0,b=0),xaxis_title=None,yaxis_title=None,coloraxis_showscale=False,paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
def export(df_,name): st.download_button('Export CSV',df_.to_csv(index=False).encode('utf-8-sig'),file_name=name,mime='text/csv')
def excel(data):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine='openpyxl') as w:
        for n,d in data.items(): d.to_excel(w,sheet_name=n[:31],index=False)
    return b.getvalue()
def header(title,sub,filters=False):
    a,b=st.columns([1.4,1])
    with a: st.title(title); st.caption(sub)
    if filters:
        with b:
            c,d=st.columns(2)
            with c: st.selectbox('Project',['All Projects']+projects,key='filter_project')
            with d: st.selectbox('Team',['All Teams']+teams,key='filter_team')
    st.markdown('---')
def filtered(data,project_col='project'):
    x=data.copy(); p=st.session_state.get('filter_project','All Projects'); t=st.session_state.get('filter_team','All Teams')
    if p!='All Projects' and project_col in x: x=x[x[project_col]==p]
    if t!='All Teams' and 'team' in x: x=x[x.team==t]
    return x

def project_form():
    with st.form('new_project'):
        st.subheader('Create Project'); name=st.text_input('Project name *'); a,b,c=st.columns(3)
        with a: owner=st.text_input('Owner')
        with b: team=st.selectbox('Team',teams)
        with c: status=st.selectbox('Status',['Not Started','On Track','At Risk','Delayed'])
        a,b=st.columns(2)
        with a: start=st.date_input('Start date',date.today())
        with b: end=st.date_input('End date',date.today()+timedelta(days=90))
        desc=st.text_area('Description'); save=st.form_submit_button('Create Project',type='primary')
        if save:
            if not name.strip() or end<start: st.error('Enter a project name and valid date range.')
            else:
                health='Critical' if status=='Delayed' else ('At Risk' if status=='At Risk' else 'Healthy')
                sql('INSERT INTO projects(name,owner,team,status,start_date,end_date,progress,description,health) VALUES(?,?,?,?,?,?,?,?,?)',(name,owner,team,status,start.isoformat(),end.isoformat(),0,desc,health)); log('Created','Project',None,name); st.success('Project created.'); st.rerun()

with st.sidebar:
    st.markdown('<div style="font-size:29px;font-weight:800;letter-spacing:-1.5px">Vektor <span style="color:#00A879">❯❯</span></div><div style="font-size:11px;margin-bottom:24px">Direction + Velocity</div>',unsafe_allow_html=True)
    nav={'▦  Overview':'Overview','▤  Projects':'Projects','▦  Gantt Chart':'Gantt Chart','☷  Task List':'Task List','♙  Resources':'Resources','▦  Calendar':'Calendar','▥  Reports':'Reports','⚙  Settings':'Settings'}
    if 'page' not in st.session_state: st.session_state.page='Overview'
    for label,name in nav.items():
        if st.button(label,key='nav_'+name,use_container_width=True,type='primary' if st.session_state.page==name else 'secondary'):
            st.session_state.page=name; st.rerun()
    st.markdown('---'); st.caption('PROJECT WORKSPACE'); st.markdown('**Local workspace**'); st.caption('SQLite storage • Demo data')
    if st.button('＋ New Project',use_container_width=True,type='primary'): st.session_state.page='Projects'; st.session_state.open_create=True; st.rerun()
page=st.session_state.page

if page=='Overview':
    header('Overview','Your project portfolio at a glance',True); p=filtered(P,'name'); t=filtered(T); total=len(t)
    c=st.columns(4)
    for col,args in zip(c,[('Total Projects',len(p),f"{int((p.status=='On Track').sum())} on track",'📁'),('Total Tasks',total,f"{int((t.status=='Completed').sum())} completed",'☑️'),('Team Members',len(R),f"{int((R.status!='On Leave').sum())} active",'👥'),('Overall Progress',f"{int(p.progress.mean()) if len(p) else 0}%",'Portfolio average','◉')]):
        with col: metric(*args)
    a,b,c=st.columns(3)
    for col,title,data,colname in [(a,'Project Status',p,'status'),(b,'Task Status',t,'status'),(c,'Team Availability',R,'status')]:
        with col,st.container(border=True): section(title); chart_status(data,colname)
    a,b,c=st.columns(3)
    with a,st.container(border=True): section('Project Progress'); progress_chart(p.sort_values('progress'))
    with b,st.container(border=True):
        section('Task Completion Trend'); trend=pd.DataFrame({'Month':['Apr','May','Jun','Jul','Aug','Sep'],'Completed':[8,18,27,35,40,48],'In Progress':[3,12,16,19,21,24],'Not Started':[2,5,7,8,9,10]}); fig=px.line(trend,x='Month',y=['Completed','In Progress','Not Started'],markers=True,color_discrete_map={'Completed':GREEN,'In Progress':BLUE,'Not Started':GRAY}); fig.update_layout(height=260,margin=dict(l=0,r=0,t=0,b=0),paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
    with c,st.container(border=True):
        section('Team Workload'); cnt=t.groupby('assignee').size().rename('Tasks').reset_index().sort_values('Tasks',ascending=False).head(6)
        if not cnt.empty:
            fig=px.bar(cnt,x='Tasks',y='assignee',orientation='h',color='Tasks',color_continuous_scale=['#00875A','#EF4444']); fig.update_layout(height=260,margin=dict(l=0,r=0,t=0,b=0),coloraxis_showscale=False,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
    a,b,c=st.columns([1.2,1,1])
    with a,st.container(border=True): section('Upcoming Milestones'); st.dataframe(M[['name','project','due_date','status']].head(6).rename(columns={'name':'Milestone','project':'Project','due_date':'Due Date','status':'Status'}),hide_index=True,use_container_width=True)
    with b,st.container(border=True): section('Top Projects by Progress'); progress_chart(p.sort_values('progress',ascending=False).head(5))
    with c,st.container(border=True): section('Project Health'); st.dataframe(p[['name','progress','health']].rename(columns={'name':'Project','progress':'Progress %','health':'Health'}),hide_index=True,use_container_width=True)

elif page=='Projects':
    header('Projects','Plan, organize, and monitor your project portfolio',True); tabs=st.tabs(['Project Portfolio','Create Project','Edit / Update'])
    with tabs[0]:
        a,b,c=st.columns(3)
        with a: metric('Total Projects',len(P),'Portfolio size','📁')
        with b: metric('Completed Projects',int((P.progress>=100).sum()),'100% progress','✅')
        with c: metric('In Progress',int(((P.progress>0)&(P.progress<100)).sum()),'Active delivery','▶️')
        q=st.text_input('Search projects'); v=P.copy()
        if q: v=v[v.astype(str).apply(lambda x:x.str.contains(q,case=False,na=False)).any(axis=1)]
        st.dataframe(v[['id','name','owner','team','status','start_date','end_date','progress','health']],hide_index=True,use_container_width=True); export(v,'projects.csv')
    with tabs[1]: project_form()
    with tabs[2]:
        if P.empty: st.info('Create a project first.')
        else:
            opts={f"{r['name']} (#{r['id']})":int(r.id) for _,r in P.iterrows()}; chosen=st.selectbox('Select project',list(opts)); pid=opts[chosen]; row=P[P.id==pid].iloc[0]
            with st.form('edit_project'):
                name=st.text_input('Project name',row['name']); a,b,c=st.columns(3)
                with a: owner=st.text_input('Owner',row.owner or '')
                with b: team=st.selectbox('Team',teams,index=teams.index(row.team) if row.team in teams else 0)
                with c: status=st.selectbox('Status',['Not Started','On Track','At Risk','Delayed'],index=['Not Started','On Track','At Risk','Delayed'].index(row.status) if row.status in ['Not Started','On Track','At Risk','Delayed'] else 0)
                a,b=st.columns(2)
                with a: start=st.date_input('Start date',date.fromisoformat(row.start_date))
                with b: end=st.date_input('End date',date.fromisoformat(row.end_date))
                progress=st.slider('Progress (%)',0,100,int(row.progress)); desc=st.text_area('Description',row.description or ''); save=st.form_submit_button('Save Changes',type='primary')
                if save:
                    if not name.strip() or end<start: st.error('Enter a valid name and date range.')
                    else:
                        health='Critical' if status=='Delayed' else ('At Risk' if status=='At Risk' or progress<50 else 'Healthy')
                        sql('UPDATE projects SET name=?,owner=?,team=?,status=?,start_date=?,end_date=?,progress=?,description=?,health=? WHERE id=?',(name,owner,team,status,start.isoformat(),end.isoformat(),progress,desc,health,pid)); log('Updated','Project',pid,name); st.success('Project updated.'); st.rerun()
            with st.expander('Delete project and linked records'):
                if st.button('Delete selected project'):
                    sql('DELETE FROM tasks WHERE project_id=?',(pid,)); sql('DELETE FROM milestones WHERE project_id=?',(pid,)); sql('DELETE FROM projects WHERE id=?',(pid,)); log('Deleted','Project',pid,chosen); st.rerun()

elif page=='Gantt Chart':
    header('Gantt Chart','Visualize schedules, task durations, dependencies, and milestones',True); t=filtered(T)
    if t.empty: st.info('Add tasks to display the Gantt chart.')
    else:
        g=t.copy(); g['Start']=pd.to_datetime(g.start_date,errors='coerce'); g['Finish']=pd.to_datetime(g.due_date,errors='coerce'); g=g.dropna(subset=['Start','Finish'])
        if not g.empty:
            fig=px.timeline(g,x_start='Start',x_end='Finish',y='name',color='status',hover_data=['project','assignee','priority','progress'],color_discrete_map={'Completed':GREEN,'In Progress':BLUE,'Not Started':GRAY,'Overdue':RED}); fig.update_yaxes(autorange='reversed',title=None); fig.update_layout(height=max(360,42*len(g)),paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        st.caption('Task dependencies are stored and displayed; automatic resource leveling and dependency-line routing are not included in this starter build.')
        st.dataframe(t[['name','project','start_date','due_date','dependency','status','progress']],hide_index=True,use_container_width=True)

elif page=='Task List':
    header('Task List','Manage task scope, assignments, priorities, and delivery dates',True); t=filtered(T); a,b,c,d=st.columns(4)
    for col,args in zip([a,b,c,d],[('Total Tasks',len(t),'Current view','☑️'),('Completed',int((t.status=='Completed').sum()),'Completed tasks','✅'),('In Progress',int((t.status=='In Progress').sum()),'Active tasks','▶️'),('Not Started',int((t.status=='Not Started').sum()),'Queued tasks','⏸️')]):
        with col: metric(*args)
    ta,tb=st.tabs(['Task Table','Create / Edit Task'])
    with ta:
        q=st.text_input('Search tasks'); v=t.copy()
        if q: v=v[v.astype(str).apply(lambda x:x.str.contains(q,case=False,na=False)).any(axis=1)]
        st.dataframe(v[['id','name','project','team','status','priority','start_date','due_date','assignee','progress']],hide_index=True,use_container_width=True); export(v,'tasks.csv')
    with tb:
        mode=st.radio('Action',['Create task','Edit task'],horizontal=True); row=None
        if mode=='Edit task' and not T.empty:
            opts={f"{r['name']} (#{r['id']})":int(r.id) for _,r in T.iterrows()}; tid=opts[st.selectbox('Select task',list(opts))]; row=T[T.id==tid].iloc[0]
        with st.form('task_form'):
            name=st.text_input('Task name *',str(row['name']) if row is not None else ''); a,b,c=st.columns(3)
            with a:
                proj_default=row.project if row is not None else (projects[0] if projects else '')
                proj=st.selectbox('Project',projects,index=projects.index(proj_default) if proj_default in projects else 0)
            with b:
                td=str(row.team) if row is not None else teams[0]; team=st.selectbox('Team',teams,index=teams.index(td) if td in teams else 0)
            with c:
                statuses=['Not Started','In Progress','Completed','Overdue']; sd=str(row.status) if row is not None else 'Not Started'; status=st.selectbox('Status',statuses,index=statuses.index(sd) if sd in statuses else 0)
            a,b,c=st.columns(3)
            with a:
                priorities=['High','Medium','Low']; pr=str(row.priority) if row is not None else 'Medium'; priority=st.selectbox('Priority',priorities,index=priorities.index(pr) if pr in priorities else 1)
            with b: start=st.date_input('Start date',date.fromisoformat(row.start_date) if row is not None and row.start_date else date.today())
            with c: due=st.date_input('Due date',date.fromisoformat(row.due_date) if row is not None and row.due_date else date.today()+timedelta(days=7))
            assignees=['Unassigned']+R.name.tolist(); ad=str(row.assignee) if row is not None else 'Unassigned'; assignee=st.selectbox('Assignee',assignees,index=assignees.index(ad) if ad in assignees else 0)
            progress=st.slider('Progress',0,100,int(row.progress or 0) if row is not None else 0); dependency=st.text_input('Dependency (task name)',str(row.dependency) if row is not None else ''); notes=st.text_area('Notes',str(row.notes) if row is not None else '')
            save=st.form_submit_button('Save Task',type='primary')
            if save:
                if not name.strip() or due<start: st.error('Task name and valid dates are required.')
                else:
                    vals=(name,pmap.get(proj),team,status,priority,start.isoformat(),due.isoformat(),'' if assignee=='Unassigned' else assignee,progress,dependency,notes)
                    if mode=='Create task': sql('INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)',vals); log('Created','Task',None,name)
                    else: sql('UPDATE tasks SET name=?,project_id=?,team=?,status=?,priority=?,start_date=?,due_date=?,assignee=?,progress=?,dependency=?,notes=? WHERE id=?',vals+(int(row.id),)); log('Updated','Task',int(row.id),name)
                    st.success('Task saved.'); st.rerun()
        if mode=='Edit task' and row is not None and st.button('Delete selected task'):
            sql('DELETE FROM tasks WHERE id=?',(int(row.id),)); log('Deleted','Task',int(row.id),row['name']); st.rerun()

elif page=='Resources':
    header('Resources','Manage team members, capacity, availability, and assignments',True); a,b,c,d=st.columns(4)
    for col,args in zip([a,b,c,d],[('Total Resources',len(R),'Team directory','👥'),('Available',int((R.status=='Available').sum()),'Available now','🟢'),('Partial Allocation',int((R.status=='Partial').sum()),'Reduced capacity','🟠'),('On Leave',int((R.status=='On Leave').sum()),'Unavailable','🔴')]):
        with col: metric(*args)
    a,b=st.columns(2)
    with a,st.container(border=True): section('Resource Status'); chart_status(R,'status')
    with b,st.container(border=True):
        section('Assignments by Project'); cnt=T.groupby('project').size().rename('Assignments').reset_index()
        if not cnt.empty:
            fig=px.bar(cnt,x='Assignments',y='project',orientation='h',color='Assignments'); fig.update_layout(height=270,margin=dict(l=0,r=0,t=0,b=0),coloraxis_showscale=False,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
    ta,tb=st.tabs(['Resource Directory','Add / Edit Resource'])
    with ta:
        q=st.text_input('Search resources'); v=R.copy()
        if q: v=v[v.astype(str).apply(lambda x:x.str.contains(q,case=False,na=False)).any(axis=1)]
        v['Active Tasks']=v.name.map(T.groupby('assignee').size().to_dict()).fillna(0).astype(int); st.dataframe(v,hide_index=True,use_container_width=True); export(v,'resources.csv')
    with tb:
        mode=st.radio('Resource action',['Add resource','Edit resource'],horizontal=True); row=None
        if mode=='Edit resource' and not R.empty:
            opts={f"{r['name']} (#{r['id']})":int(r.id) for _,r in R.iterrows()}; rid=opts[st.selectbox('Select resource',list(opts))]; row=R[R.id==rid].iloc[0]
        with st.form('resource_form'):
            name=st.text_input('Full name *',str(row['name']) if row is not None else ''); email=st.text_input('Email',str(row.email) if row is not None else ''); a,b,c=st.columns(3)
            with a: td=str(row.team) if row is not None else teams[0]; team=st.selectbox('Team',teams,index=teams.index(td) if td in teams else 0)
            with b:
                roles=['Admin','Project Manager','Team Member']; rd=str(row.role) if row is not None else 'Team Member'; role=st.selectbox('Role',roles,index=roles.index(rd) if rd in roles else 2)
            with c:
                states=['Available','Partial','On Leave']; sd=str(row.status) if row is not None else 'Available'; status=st.selectbox('Availability',states,index=states.index(sd) if sd in states else 0)
            capacity=st.slider('Capacity (%)',0,100,int(row.capacity or 100) if row is not None else 100); notes=st.text_area('Notes',str(row.notes) if row is not None else ''); save=st.form_submit_button('Save Resource',type='primary')
            if save:
                if not name.strip(): st.error('Name is required.')
                elif mode=='Add resource': sql('INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)',(name,email,team,role,status,capacity,notes)); log('Created','Resource',None,name); st.rerun()
                else: sql('UPDATE resources SET name=?,email=?,team=?,role=?,status=?,capacity=?,notes=? WHERE id=?',(name,email,team,role,status,capacity,notes,int(row.id))); log('Updated','Resource',int(row.id),name); st.rerun()
        if mode=='Edit resource' and row is not None and st.button('Delete selected resource'): sql('DELETE FROM resources WHERE id=?',(int(row.id),)); log('Deleted','Resource',int(row.id)); st.rerun()

elif page=='Calendar':
    header('Calendar','View project dates, task due dates, and milestone deadlines',True); month=st.date_input('Calendar month',date.today().replace(day=1)); first=month.replace(day=1); last=month.replace(day=calendar.monthrange(month.year,month.month)[1]); events=[]
    for _,r in T.iterrows():
        try: d=date.fromisoformat(str(r.due_date))
        except: continue
        if first<=d<=last: events.append({'Date':d,'Type':'Task','Title':r['name'],'Project':r['project'],'Status':r.status})
    for _,r in M.iterrows():
        try: d=date.fromisoformat(str(r.due_date))
        except: continue
        if first<=d<=last: events.append({'Date':d,'Type':'Milestone','Title':r['name'],'Project':r['project'],'Status':r.status})
    ev=pd.DataFrame(events); st.subheader(first.strftime('%B %Y')); weeks=calendar.Calendar(firstweekday=6).monthdatescalendar(first.year,first.month)
    for col,dayname in zip(st.columns(7),['Sun','Mon','Tue','Wed','Thu','Fri','Sat']): col.markdown(f'**{dayname}**')
    for week in weeks:
        cols=st.columns(7)
        for col,d in zip(cols,week):
            with col:
                st.markdown(f'**{d.day}**' if d.month==first.month else f'<span style="color:#aab5c4">{d.day}</span>',unsafe_allow_html=True)
                daily=ev[ev.Date==d] if not ev.empty else pd.DataFrame()
                for _,e in daily.head(3).iterrows():
                    color=GREEN if e['Type']=='Milestone' else BLUE; st.markdown(f'<div style="background:{color};color:white;padding:3px;border-radius:4px;font-size:10px;margin:2px 0">{e["Type"]}: {e["Title"][:22]}</div>',unsafe_allow_html=True)
    section('Events in Selected Month')
    if ev.empty: st.info('No events due in this month.')
    else: st.dataframe(ev.sort_values('Date'),hide_index=True,use_container_width=True)

elif page=='Reports':
    st.title('Reports'); st.caption('Gain insights into project progress, resource utilization, and team performance.')
    a,b,c,d=st.columns([1.25,1,1,1])
    with a: dates=st.date_input('Date Range',(date(2024,4,24),date(2024,6,30)),key='report_dates')
    with b: pf=st.selectbox('Project',['All Projects']+projects,key='report_project')
    with c: tf=st.selectbox('Team',['All Teams']+teams,key='report_team')
    with d: fmt=st.selectbox('Export format',['Excel (.xlsx)','CSV (.csv)','JSON (.json)'])
    tabs=st.tabs(['Overview','Projects','Resources','Tasks','Milestones','Workload','Custom'])
    rp=P.copy(); rt=T.copy(); rr=R.copy(); rm=M.copy()
    if pf!='All Projects': rp=rp[rp.name==pf]; rt=rt[rt.project==pf]; rm=rm[rm.project==pf]
    if tf!='All Teams': rp=rp[rp.team==tf]; rt=rt[rt.team==tf]; rr=rr[rr.team==tf]; rm=rm[rm.team==tf]
    if isinstance(dates,tuple) and len(dates)==2:
        ds,de=map(lambda x:x.isoformat(),dates); rt=rt[(rt.start_date.fillna('')<=de)&(rt.due_date.fillna('')>=ds)]; rm=rm[(rm.due_date.fillna('')>=ds)&(rm.due_date.fillna('')<=de)]
    def report_export(frames,stem):
        if fmt.startswith('Excel'): st.download_button('⬇ Export Report',excel(frames),file_name=stem+'.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='exp_'+stem)
        elif fmt.startswith('CSV'): st.download_button('⬇ Export Report',next(iter(frames.values())).to_csv(index=False).encode('utf-8-sig'),file_name=stem+'.csv',mime='text/csv',key='exp_'+stem)
        else: st.download_button('⬇ Export Report',json.dumps({k:v.where(pd.notna(v),None).to_dict('records') for k,v in frames.items()},default=str,indent=2).encode(),file_name=stem+'.json',mime='application/json',key='exp_'+stem)
    with tabs[0]:
        cols=st.columns(4)
        for col,args in zip(cols,[('Total Projects',len(rp),'','📁'),('Total Tasks',len(rt),'','☑️'),('Milestones',len(rm),'','⚑'),('Team Members',len(rr),'','👥')]):
            with col: metric(*args)
        a,b,c=st.columns(3)
        for col,title,data,field in [(a,'Project Status',rp,'status'),(b,'Task Status',rt,'status'),(c,'Resource Availability',rr,'status')]:
            with col,st.container(border=True): section(title); chart_status(data,field)
        a,b=st.columns(2)
        with a,st.container(border=True): section('Project Progress'); progress_chart(rp.sort_values('progress'))
        with b,st.container(border=True):
            section('Task Completion Trend'); trend=pd.DataFrame({'Month':['Apr','May','Jun','Jul','Aug','Sep'],'Completed':[8,18,27,35,40,48],'In Progress':[3,12,16,19,21,24]}); fig=px.line(trend,x='Month',y=['Completed','In Progress'],markers=True,color_discrete_map={'Completed':GREEN,'In Progress':BLUE}); fig.update_layout(height=260,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        report_export({'Projects':rp,'Tasks':rt,'Resources':rr,'Milestones':rm},'vektor_overview_report')
    with tabs[1]:
        st.subheader('Project Report'); cols=st.columns(4)
        for col,args in zip(cols,[('Total Projects',len(rp),'','📁'),('Completed',int((rp.progress>=100).sum()),'','✅'),('In Progress',int(((rp.progress>0)&(rp.progress<100)).sum()),'','▶️'),('At Risk / Delayed',int(rp.status.isin(['At Risk','Delayed']).sum()),'','⚠️')]):
            with col: metric(*args)
        a,b=st.columns(2)
        with a,st.container(border=True): section('Project Status'); chart_status(rp,'status')
        with b,st.container(border=True): section('Project Progress'); progress_chart(rp.sort_values('progress'))
        st.dataframe(rp,hide_index=True,use_container_width=True); report_export({'Projects':rp},'vektor_project_report')
    with tabs[2]:
        st.subheader('Resource Report'); cols=st.columns(4)
        for col,args in zip(cols,[('Total Resources',len(rr),'','👥'),('Available',int((rr.status=='Available').sum()),'','🟢'),('Partial',int((rr.status=='Partial').sum()),'','🟠'),('On Leave',int((rr.status=='On Leave').sum()),'','🔴')]):
            with col: metric(*args)
        a,b=st.columns(2)
        with a,st.container(border=True): section('Resource Status'); chart_status(rr,'status')
        with b,st.container(border=True):
            section('Team Distribution'); td=rr.team.value_counts().rename_axis('Team').reset_index(name='Members'); fig=px.bar(td,x='Team',y='Members',color='Team'); fig.update_layout(height=260,showlegend=False,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(rr,hide_index=True,use_container_width=True); report_export({'Resources':rr},'vektor_resource_report')
    with tabs[3]:
        st.subheader('Tasks Report'); cols=st.columns(4)
        for col,args in zip(cols,[('Total Tasks',len(rt),'','☑️'),('Completed',int((rt.status=='Completed').sum()),'','✅'),('In Progress',int((rt.status=='In Progress').sum()),'','▶️'),('Not Started',int((rt.status=='Not Started').sum()),'','⏸️')]):
            with col: metric(*args)
        a,b,c=st.columns(3)
        with a,st.container(border=True): section('Task Status'); chart_status(rt,'status')
        with b,st.container(border=True):
            section('Tasks by Project'); z=rt.groupby('project').size().rename('Tasks').reset_index(); fig=px.bar(z,x='Tasks',y='project',orientation='h'); fig.update_layout(height=270,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        with c,st.container(border=True):
            section('Tasks by Priority'); z=rt.priority.value_counts().rename_axis('Priority').reset_index(name='Tasks'); fig=px.pie(z,names='Priority',values='Tasks',hole=.5,color='Priority',color_discrete_map={'High':RED,'Medium':ORANGE,'Low':GREEN}); fig.update_layout(height=270,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(rt,hide_index=True,use_container_width=True); report_export({'Tasks':rt},'vektor_task_report')
    with tabs[4]:
        st.subheader('Milestones Report'); overdue=rm[(pd.to_datetime(rm.due_date,errors='coerce').dt.date<date.today())&(rm.status!='Completed')] if not rm.empty else rm; cols=st.columns(4)
        for col,args in zip(cols,[('Total Milestones',len(rm),'','⚑'),('Completed',int((rm.status=='Completed').sum()),'','✅'),('In Progress',int((rm.status=='In Progress').sum()),'','▶️'),('Overdue',len(overdue),'','⚠️')]):
            with col: metric(*args)
        a,b=st.columns(2)
        with a,st.container(border=True): section('Milestone Status'); chart_status(rm,'status')
        with b,st.container(border=True):
            section('Milestones by Project'); z=rm.groupby('project').size().rename('Milestones').reset_index(); fig=px.bar(z,x='project',y='Milestones',color='Milestones'); fig.update_layout(height=260,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(rm,hide_index=True,use_container_width=True); report_export({'Milestones':rm},'vektor_milestone_report')
    with tabs[5]:
        st.subheader('Workload Report'); counts=rt.groupby('assignee').size().rename('Active Tasks').reset_index() if not rt.empty else pd.DataFrame(columns=['assignee','Active Tasks']); work=rr.merge(counts,left_on='name',right_on='assignee',how='left').drop(columns=['assignee'],errors='ignore'); work['Active Tasks']=work['Active Tasks'].fillna(0).astype(int); work['Workload Status']=work['Active Tasks'].apply(lambda x:'Overloaded' if x>=15 else ('High' if x>=10 else 'Balanced'))
        cols=st.columns(4)
        for col,args in zip(cols,[('Team Members',len(work),'','👥'),('Active Tasks',int(work['Active Tasks'].sum()),'','☑️'),('High Workload',int((work['Workload Status']=='High').sum()),'','🟠'),('Overloaded',int((work['Workload Status']=='Overloaded').sum()),'','🔴')]):
            with col: metric(*args)
        a,b=st.columns(2)
        with a,st.container(border=True):
            section('Individual Workload'); z=work.sort_values('Active Tasks',ascending=False).head(10); fig=px.bar(z,x='Active Tasks',y='name',orientation='h',color='Workload Status',color_discrete_map={'Balanced':GREEN,'High':ORANGE,'Overloaded':RED}); fig.update_layout(height=320,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        with b,st.container(border=True):
            section('Workload by Team'); z=work.groupby('team')['Active Tasks'].sum().reset_index(); fig=px.bar(z,x='team',y='Active Tasks',color='team'); fig.update_layout(height=320,showlegend=False,paper_bgcolor='rgba(0,0,0,0)'); st.plotly_chart(fig,use_container_width=True)
        st.dataframe(work,hide_index=True,use_container_width=True); report_export({'Workload':work},'vektor_workload_report')
    with tabs[6]:
        st.subheader('Build Custom Report')
        with st.container(border=True):
            a,b,c=st.columns(3)
            with a: typ=st.selectbox('Report Type',['Project Report','Task Report','Resource Report','Milestone Report','Workload Report'])
            with b: group=st.selectbox('Group By',['Project','Team','Status','Priority','Assignee'])
            with c: metrics=st.multiselect('Metrics / Data to Include',['Project Status','Task Status','Milestone Status','Team Members','Progress','Priority','Dates','Assignees'],default=['Project Status','Task Status','Milestone Status'])
            viz=st.radio('Visualization Type',['Bar Chart','Column Chart','Pie Chart','Line Chart','Table','Summary Cards'],horizontal=True)
            a,b=st.columns(2)
            with a: generate=st.button('▶ Generate Report',type='primary')
            with b: reset=st.button('Reset')
            if reset: st.session_state.pop('custom_report',None)
            if generate: st.session_state.custom_report={'type':typ,'group':group,'metrics':metrics,'viz':viz}
        if st.session_state.get('custom_report'):
            cfg=st.session_state.custom_report; source={'Project Report':rp,'Task Report':rt,'Resource Report':rr,'Milestone Report':rm}.get(cfg['type'],rr)
            colmap={'Project':'project','Team':'team','Status':'status','Priority':'priority','Assignee':'assignee'}; gc=colmap[cfg['group']]
            if gc not in source: gc='name' if 'name' in source else source.columns[0]
            z=source.groupby(gc).size().rename('Count').reset_index(); st.subheader('Report Preview')
            if cfg['viz']=='Pie Chart': fig=px.pie(z,names=gc,values='Count',hole=.4); st.plotly_chart(fig,use_container_width=True)
            elif cfg['viz'] in ['Bar Chart','Column Chart','Line Chart']:
                fig=px.line(z,x=gc,y='Count',markers=True) if cfg['viz']=='Line Chart' else px.bar(z,x='Count' if cfg['viz']=='Bar Chart' else gc,y=gc if cfg['viz']=='Bar Chart' else 'Count',orientation='h' if cfg['viz']=='Bar Chart' else 'v'); st.plotly_chart(fig,use_container_width=True)
            elif cfg['viz']=='Summary Cards':
                cols=st.columns(min(4,max(1,len(z))))
                for col,(_,r) in zip(cols,z.iterrows()):
                    with col: metric(str(r[gc]),int(r.Count),'records','📊')
            st.dataframe(source,hide_index=True,use_container_width=True); report_export({cfg['type'].split()[0]:source},'vektor_custom_report')
            with st.form('save_report'):
                report_name=st.text_input('Save report as',f"{cfg['type']} - {date.today().isoformat()}"); save=st.form_submit_button('＋ Save Current Report')
                if save: sql('INSERT INTO saved_reports(name,report_type,config,last_generated,created_by) VALUES(?,?,?,?,?)',(report_name,cfg['type'],json.dumps(cfg),datetime.now().isoformat(timespec='seconds'),'Admin')); log('Saved','Report',None,report_name); st.success('Report saved.')
        with st.expander('Saved Reports'):
            saved=df('SELECT * FROM saved_reports ORDER BY id DESC')
            if saved.empty: st.caption('No saved reports yet.')
            else:
                st.dataframe(saved,hide_index=True,use_container_width=True); rid=st.selectbox('Select saved report',saved.id.tolist());
                if st.button('Delete Saved Report'): sql('DELETE FROM saved_reports WHERE id=?',(int(rid),)); st.rerun()
                export(saved,'saved_reports.csv')

elif page=='Settings':
    header('Settings','Manage workspace, users, defaults, notifications, integrations, and system preferences'); tabs=st.tabs(['Workspace','User Management','Project Settings','Task Settings','Milestones','Notifications','Integrations','System'])
    def get(k,default=''):
        x=df('SELECT value FROM settings WHERE key=?',(k,)); return str(x.iloc[0].value) if not x.empty else default
    def put(k,v): sql('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',(k,str(v)))
    with tabs[0]:
        with st.form('workspace'):
            st.subheader('Workspace Settings'); name=st.text_input('Workspace Name',get('workspace_name','Vektor')); desc=st.text_area('Description',get('workspace_description','Project management workspace for planning, tracking, and delivery.')); a,b=st.columns(2)
            with a: theme=st.selectbox('Theme',['HPE Green','Blue','Dark'],index=['HPE Green','Blue','Dark'].index(get('theme','HPE Green')) if get('theme','HPE Green') in ['HPE Green','Blue','Dark'] else 0)
            with b: language=st.selectbox('Language',['English','Filipino'])
            a,b=st.columns(2)
            with a: fmt=st.selectbox('Date Format',['MMM DD, YYYY','DD/MM/YYYY','MM/DD/YYYY'])
            with b: landing=st.selectbox('Default Landing Page',['Overview','Projects','Task List','Reports'])
            st.caption('Logo upload and global visual theme switching can be added for deployment.'); save=st.form_submit_button('Save Workspace Settings',type='primary')
            if save:
                for k,v in {'workspace_name':name,'workspace_description':desc,'theme':theme,'language':language,'date_format':fmt,'default_landing':landing}.items(): put(k,v)
                st.success('Workspace settings saved.')
    with tabs[1]:
        st.subheader('User Management'); st.caption('User records use the Resources table. Login, password management, and SSO require an authentication service.'); st.dataframe(R[['id','name','email','role','team','status']],hide_index=True,use_container_width=True)
        with st.form('add_user'):
            name=st.text_input('Name'); email=st.text_input('Email'); a,b=st.columns(2)
            with a: role=st.selectbox('Role',['Admin','Project Manager','Team Member'])
            with b: team=st.selectbox('Team',teams)
            save=st.form_submit_button('Add User',type='primary')
            if save:
                if name.strip(): sql('INSERT INTO resources(name,email,team,role,status,capacity) VALUES(?,?,?,?,?,?)',(name,email,team,role,'Available',100)); log('Created','User',None,name); st.rerun()
                else: st.error('Name is required.')
        with st.expander('Update or remove user'):
            if not R.empty:
                uid=st.selectbox('User ID',R.id.tolist(),format_func=lambda x:R.loc[R.id==x,'name'].iloc[0]); role=st.selectbox('Role',['Admin','Project Manager','Team Member'],key='role_update'); status=st.selectbox('Status',['Available','Partial','On Leave'],key='status_update'); a,b=st.columns(2)
                with a:
                    if st.button('Update User'): sql('UPDATE resources SET role=?,status=? WHERE id=?',(role,status,int(uid))); log('Updated','User',int(uid)); st.rerun()
                with b:
                    if st.button('Remove User'): sql('DELETE FROM resources WHERE id=?',(int(uid),)); log('Deleted','User',int(uid)); st.rerun()
    with tabs[2]:
        with st.form('project_defaults'):
            st.subheader('Project Settings'); status=st.selectbox('Default Project Status',['Not Started','On Track','At Risk','Delayed']); template=st.selectbox('Default Project Template',['Standard Project','Software Delivery','Marketing Campaign','Infrastructure'])
            vals={k:st.checkbox(label,value=get(k,'true' if default else 'false')=='true') for k,label,default in [('auto_project_id','Enable project key/ID auto-generation',True),('require_project_desc','Require project description',True),('allow_project_attachments','Allow project attachments',True),('project_archiving','Enable project archiving',False)]}
            if st.form_submit_button('Save Project Defaults',type='primary'):
                put('default_project_status',status); put('project_template',template)
                for k,v in vals.items(): put(k,str(v).lower())
                st.success('Project defaults saved.')
    with tabs[3]:
        with st.form('task_defaults'):
            st.subheader('Task Settings'); status=st.selectbox('Default Task Status',['Not Started','In Progress','Completed']); priority=st.selectbox('Default Task Priority',['High','Medium','Low'],index=1)
            vals={k:st.checkbox(label,value=get(k,'true' if default else 'false')=='true') for k,label,default in [('task_dependencies','Enable task dependencies',True),('allow_subtasks','Allow subtasks',True),('recurring_tasks','Enable recurring tasks',False),('require_assignee','Require assignee',True),('task_attachments','Allow task attachments',True)]}
            if st.form_submit_button('Save Task Defaults',type='primary'):
                put('default_task_status',status); put('default_task_priority',priority)
                for k,v in vals.items(): put(k,str(v).lower())
                st.success('Task defaults saved.')
    with tabs[4]:
        with st.form('milestone_defaults'):
            st.subheader('Milestone Settings'); status=st.selectbox('Default Milestone Status',['Not Started','In Progress','Completed']); vals={k:st.checkbox(label,value=get(k,'true')=='true') for k,label in [('milestones_enabled','Enable milestones'),('milestones_require_due','Require due date'),('milestone_dependencies','Allow milestone dependencies'),('milestone_notifications','Auto-notify on milestone updates')]}
            if st.form_submit_button('Save Milestone Defaults',type='primary'):
                put('default_milestone_status',status)
                for k,v in vals.items(): put(k,str(v).lower())
                st.success('Milestone defaults saved.')
        st.subheader('Add Milestone')
        with st.form('add_milestone'):
            name=st.text_input('Milestone name'); project=st.selectbox('Project',projects); due=st.date_input('Due date',date.today()+timedelta(days=14)); team=st.selectbox('Team',teams); status=st.selectbox('Status',['Not Started','In Progress','Completed','Overdue']); save=st.form_submit_button('Add Milestone')
            if save and name.strip(): sql('INSERT INTO milestones(name,project_id,team,status,start_date,due_date,notes) VALUES(?,?,?,?,?,?,?)',(name,pmap[project],team,status,date.today().isoformat(),due.isoformat(),'')); log('Created','Milestone',None,name); st.rerun()
        st.dataframe(M[['id','name','project','status','due_date']],hide_index=True,use_container_width=True)
    with tabs[5]:
        with st.form('notifications'):
            st.subheader('Notifications'); labels=['Task assigned to me','Task due soon','Task overdue','Milestone updates','Project status changes','Project assignment','Mention in comments','Weekly project digest']
            vals={k:st.checkbox(k,value=get('notify_'+k,'false' if k=='Mention in comments' else 'true')=='true') for k in labels}; freq=st.selectbox('Notification Frequency',['Real-time','Hourly','Daily digest','Weekly digest'])
            if st.form_submit_button('Save Notification Preferences',type='primary'):
                for k,v in vals.items(): put('notify_'+k,str(v).lower())
                put('notification_frequency',freq); st.success('Preferences saved. Actual message delivery requires a configured integration.')
    with tabs[6]:
        st.subheader('Integrations'); st.caption('Connect buttons are placeholders until OAuth credentials and callback URLs are configured.')
        for service,desc in [('Microsoft 365','Connect calendar and Teams'),('Slack','Receive notifications in Slack'),('Microsoft Teams','Team collaboration'),('OneDrive','Access project files')]:
            with st.container(border=True):
                a,b=st.columns([3,1])
                with a: st.markdown('**'+service+'**'); st.caption(desc)
                with b:
                    if st.button('Connect',key='connect_'+service): st.info(f'{service} integration is not configured. Configure OAuth credentials before enabling.')
    with tabs[7]:
        with st.form('system_settings'):
            st.subheader('System Settings'); audit=st.checkbox('Enable audit logs',value=get('audit_logs','true')=='true'); exports=st.checkbox('Allow data export',value=get('data_export','true')=='true'); backups=st.checkbox('Enable automatic backups',value=get('automatic_backups','true')=='true'); maintenance=st.checkbox('Enable maintenance mode',value=get('maintenance_mode','false')=='true'); timeout=st.selectbox('Session Timeout',['30 minutes','1 hour','4 hours','8 hours','12 hours'],index=3)
            if st.form_submit_button('Save System Settings',type='primary'):
                for k,v in {'audit_logs':audit,'data_export':exports,'automatic_backups':backups,'maintenance_mode':maintenance,'session_timeout':timeout}.items(): put(k,str(v).lower() if isinstance(v,bool) else v)
                st.success('System settings saved.')
        st.subheader('Data Management'); a,b=st.columns(2)
        with a:
            if st.button('Prepare Database Backup'):
                with open(DB,'rb') as f: data=f.read()
                st.download_button('Download SQLite Backup',data,file_name=f'vektor_backup_{date.today().isoformat()}.db',mime='application/octet-stream')
        with b:
            if st.button('Clear Audit Log'): sql('DELETE FROM audit_log'); st.success('Audit log cleared.')
        with st.expander('Recent Audit Events'):
            logs=df('SELECT * FROM audit_log ORDER BY id DESC LIMIT 100'); st.dataframe(logs,hide_index=True,use_container_width=True)

st.markdown('---')
a,b=st.columns([3,1])
with a: st.caption(f'Vektor Project Management • Local SQLite workspace • Database: {DB}')
with b: st.caption('Updated '+datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
