from datetime import date, datetime, timedelta
import calendar, json, uuid
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Vektor | Project Management', page_icon='◆', layout='wide')
st.markdown('''<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:Inter,sans-serif}.stApp{background:#edf4f9;color:#101d49}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#061c22,#00191e)}[data-testid="stSidebar"] *{color:#f3f8fb!important}
.brand{font-size:30px;font-weight:800;color:#f7fbff;padding-top:10px}.brand span{color:#00a878}.sub{color:#c2d3d9;font-size:11px;margin-bottom:22px}
.h1{font-size:30px;font-weight:800;letter-spacing:-1px;color:#111d4d}.subhead{color:#5b6b88;margin-bottom:18px}
.metric{background:white;border:1px solid #dce7f0;border-radius:12px;padding:15px;min-height:90px;box-shadow:0 2px 8px #132c5010}.ml{font-size:12px;color:#5c6b85;font-weight:600}.mv{font-size:25px;font-weight:800;color:#101d49;margin-top:5px}
.stButton button{border-radius:8px;font-weight:600}
</style>''',unsafe_allow_html=True)

def uid(p): return f'{p}-{uuid.uuid4().hex[:7].upper()}'
today=date.today()
def seed():
    ps=[
      dict(id='PRJ-001',name='Website Revamp',owner='Alice Morgan',start=today-timedelta(days=20),end=today+timedelta(days=45),status='In Progress',priority='High',description='Refresh the company website experience, navigation and conversion journey.',budget=85000),
      dict(id='PRJ-002',name='IT Infrastructure Upgrade',owner='Ryan Cruz',start=today-timedelta(days=5),end=today+timedelta(days=60),status='In Progress',priority='High',description='Upgrade core infrastructure, security and endpoint tooling.',budget=120000),
      dict(id='PRJ-003',name='Sustainability Initiative',owner='Emily Rodriguez',start=today+timedelta(days=2),end=today+timedelta(days=75),status='Planning',priority='Medium',description='Identify and deliver sustainability improvements across operations.',budget=42000),
      dict(id='PRJ-004',name='Marketing Campaign',owner='Alice Morgan',start=today-timedelta(days=10),end=today+timedelta(days=30),status='In Progress',priority='Medium',description='Launch a coordinated campaign across digital channels.',budget=60000)]
    rs=[]
    raw=[('James Miller','UX Designer','UI/UX,Figma',100,['PRJ-001']),('Emily Rodriguez','Business Analyst','Analysis,Requirements',80,['PRJ-003']),('Sarah Kim','Frontend Developer','React,JavaScript',100,['PRJ-001']),('Alice Morgan','Project Manager','Planning,Stakeholder',50,['PRJ-004']),('Ryan Cruz','Backend Developer','Python,API',90,['PRJ-002']),('Michael Tan','QA Engineer','Testing,Automation',120,['PRJ-002']),('Dana Scott','Content Specialist','Content,SEO',40,['PRJ-004']),('Karen Lim','DevOps Engineer','Azure,CI/CD',100,['PRJ-002']),('Noah Park','IT Support','Infrastructure,Support',60,['PRJ-002']),('Chloe Tan','Data Analyst','Power BI,Excel',100,['PRJ-003'])]
    for i,(n,role,skills,alloc,pids) in enumerate(raw): rs.append(dict(id=f'RES-{i+1:03}',name=n,email=n.lower().replace(' ','.')+'@example.com',role=role,skills=skills.split(','),allocation=alloc,status='Overallocated' if alloc>100 else ('Available' if alloc<70 else 'Allocated'),projects=pids,hours=8,capacity=100,notes=''))
    ts=[]; names=['Define project scope','Stakeholder alignment','Create wireframes','UI/UX design','Frontend development','Backend development','UAT testing','Go-Live']; pcts=[100,100,100,70,50,20,0,0]; durations=[3,4,7,10,12,15,5,2]
    assignees=['Alice Morgan','Emily Rodriguez','James Miller','Sarah Kim','Sarah Kim','Ryan Cruz','Michael Tan','Alice Morgan']
    for i,n in enumerate(names):
        sd=today-timedelta(days=18-i*2); pct=pcts[i]
        ts.append(dict(id=f'1.{i+1}',name=n,project_id='PRJ-001',parent='',description=f'{n} for the Website Revamp project.',assignee=assignees[i],start=sd,end=sd+timedelta(days=durations[i]-1),duration=durations[i],priority='High' if i in (2,3) else 'Medium',status='Completed' if pct==100 else ('In Progress' if pct else 'Not Started'),type='Milestone' if n=='Go-Live' else 'Task',milestone=n=='Go-Live',percent=pct,tags=['Planning' if i<2 else 'Design' if i<4 else 'Development' if i<6 else 'Launch'],dependencies=[],checklist=[],estimated=durations[i]*5,actual=durations[i]*3 if pct else 0,notes='',attachments=[],activity=[]))
    for pi,p in enumerate(ps[1:],2):
        for j,n in enumerate(['Requirements','Implementation','Testing']):
            sd=p['start']+timedelta(days=j*8); pct=100 if j==0 else 40 if j==1 else 0
            ts.append(dict(id=f'{pi}.{j+1}',name=n,project_id=p['id'],parent='',description=f'{n} for {p["name"]}.',assignee=rs[(pi+j)%len(rs)]['name'],start=sd,end=sd+timedelta(days=7),duration=8,priority='Medium',status='Completed' if j==0 else 'In Progress' if j==1 else 'Not Started',type='Task',milestone=False,percent=pct,tags=[n],dependencies=[],checklist=[],estimated=40,actual=16 if j==1 else 0,notes='',attachments=[],activity=[]))
    return ps,rs,ts,[]
if 'projects' not in st.session_state: st.session_state.projects,st.session_state.resources,st.session_state.tasks,st.session_state.events=seed()
P,R,T,E=st.session_state.projects,st.session_state.resources,st.session_state.tasks,st.session_state.events
if 'page' not in st.session_state: st.session_state.page='Overview'
if 'task_action' not in st.session_state: st.session_state.task_action=None
if 'resource_action' not in st.session_state: st.session_state.resource_action=None

def pname(pid): return next((p['name'] for p in P if p['id']==pid),'Unassigned')
def metric(label,value,icon='◆'):
    st.markdown(f'<div class="metric"><div class="ml">{icon} &nbsp; {label}</div><div class="mv">{value}</div></div>',unsafe_allow_html=True)
def header(title,subtitle): st.markdown(f'<div class="h1">{title}</div><div class="subhead">{subtitle}</div>',unsafe_allow_html=True)
def task_table(data,key):
    if not data: st.info('No tasks match the selected filters.'); return
    df=pd.DataFrame([{'ID':t['id'],'Task Name':t['name'],'Project':pname(t['project_id']),'Assignee':t['assignee'],'Start Date':t['start'].strftime('%b %d, %Y'),'End Date':t['end'].strftime('%b %d, %Y'),'Duration':t['duration'],'Status':t['status'],'Priority':t['priority'],'% Complete':t['percent'],'Tags':', '.join(t['tags'])} for t in data])
    st.dataframe(df,use_container_width=True,hide_index=True)
    opts={f"{t['id']} — {t['name']}":t['id'] for t in data}; choice=st.selectbox('Select a task to view or edit',['']+list(opts),key=key+'_select')
    if choice:
        t=next(x for x in data if x['id']==opts[choice]); c1,c2,c3=st.columns([1,1,5])
        if c1.button('View',key=key+'_view'): st.session_state.task_action=('view',t['id'])
        if c2.button('Edit',key=key+'_edit'): st.session_state.task_action=('edit',t['id'])

with st.sidebar:
    st.markdown('<div class="brand">Vektor <span>»</span></div><div class="sub">Direction + Velocity</div>',unsafe_allow_html=True)
    for icon,label in [('▦','Overview'),('⊞','Projects'),('▤','Gantt Chart'),('☷','Task List'),('♙','Resources'),('▦','Calendar'),('▥','Reports'),('⚙','Settings')]:
        if st.button(f'{icon}   {label}',key='nav_'+label,use_container_width=True,type='primary' if st.session_state.page==label else 'secondary'):
            st.session_state.page=label; st.rerun()
    st.markdown('---'); st.caption(f'{len(P)} projects · {len(T)} tasks · {len(R)} resources')
page=st.session_state.page
if page=='Overview':
    header('Overview','A live snapshot of projects, tasks, and team capacity.')
    cols=st.columns(4); vals=[('Active Projects',sum(p['status']=='In Progress' for p in P)),('Total Tasks',len(T)),('At Risk',sum(t['priority']=='High' and t['status']!='Completed' for t in T)),('Overdue',sum(t['end']<today and t['status']!='Completed' for t in T))]
    for c,(l,v) in zip(cols,vals):
        with c: metric(l,v)
    c1,c2=st.columns([1.2,1]);
    with c1:
        st.subheader('Project Progress'); st.dataframe(pd.DataFrame([{'Project':p['name'],'Tasks':sum(t['project_id']==p['id'] for t in T),'Status':p['status'],'Deadline':p['end']} for p in P]),use_container_width=True,hide_index=True)
    with c2: st.subheader('Task Status'); st.bar_chart(pd.Series({s:sum(t['status']==s for t in T) for s in ['Not Started','In Progress','Completed','On Hold']}))
    st.subheader('Priority Tasks'); task_table(sorted([t for t in T if t['status']!='Completed'],key=lambda t:(t['priority'] not in ('Critical','High'),t['end']))[:8],'overview')
elif page=='Projects':
    header('Projects','Create, plan, and track all projects.')
    with st.expander('＋ Add Project'):
        with st.form('project_form'):
            name=st.text_input('Project name *'); owner=st.selectbox('Project manager',[r['name'] for r in R]); c1,c2=st.columns(2); start=c1.date_input('Start date',today); end=c2.date_input('End date',today+timedelta(days=30)); status=st.selectbox('Status',['Planning','In Progress','On Hold','Completed']); priority=st.selectbox('Priority',['Low','Medium','High','Critical']); budget=st.number_input('Budget',min_value=0.0,step=1000.0); desc=st.text_area('Description')
            if st.form_submit_button('Create Project',type='primary'):
                if not name.strip() or end<start: st.error('Enter a name and a valid date range.')
                else: P.append(dict(id=uid('PRJ'),name=name,owner=owner,start=start,end=end,status=status,priority=priority,description=desc,budget=budget)); st.success('Project created.'); st.rerun()
    q=st.text_input('Search projects')
    for p in P:
        if q.lower() not in (p['name']+p['owner']+p['status']).lower(): continue
        with st.container(border=True):
            a,b,c=st.columns([3,1,1]); a.markdown(f"### {p['name']}"); a.caption(f"{p['id']} · Owner: {p['owner']}"); a.write(p['description']); b.metric('Tasks',sum(t['project_id']==p['id'] for t in T)); c.write(p['status']); c.caption(f"{p['start']} – {p['end']}")
            if a.button('View project tasks',key='viewproj_'+p['id']): st.session_state.project_filter=p['id']; st.session_state.page='Task List'; st.rerun()
elif page=='Task List':
    header('Task List','Manage, track, and update all project tasks in one place.')
    with st.expander('＋ Add Task',expanded=False):
        with st.form('add_task_form',clear_on_submit=True):
            c1,c2=st.columns(2); project=c1.selectbox('Project *',[p['name'] for p in P]); parent=c2.selectbox('Parent Task',['None']+[t['id']+' — '+t['name'] for t in T if t['project_id']==next(p['id'] for p in P if p['name']==project)]); name=st.text_input('Task name *'); desc=st.text_area('Description'); c1,c2,c3=st.columns(3); start=c1.date_input('Start date',today); end=c2.date_input('End date',today+timedelta(days=7)); priority=c3.selectbox('Priority',['Low','Medium','High','Critical']); c1,c2,c3=st.columns(3); assignee=c1.selectbox('Assignee',[r['name'] for r in R]); status=c2.selectbox('Status',['Not Started','In Progress','On Hold','Completed']); pct=c3.slider('% Complete',0,100,0); tags=st.text_input('Tags (comma separated)'); estimated=st.number_input('Estimated effort (hours)',min_value=0,step=1); checklist=st.text_area('Checklist / subtasks (one per line)'); notes=st.text_area('Progress notes')
            if st.form_submit_button('Save Task',type='primary'):
                if not name.strip() or end<start: st.error('Enter a task name and a valid date range.')
                else:
                    pid=next(p['id'] for p in P if p['name']==project); T.append(dict(id=uid('TASK'),name=name,project_id=pid,parent='' if parent=='None' else parent.split(' — ')[0],description=desc,assignee=assignee,start=start,end=end,duration=(end-start).days+1,priority=priority,status=status,type='Task',milestone=False,percent=pct,tags=[x.strip() for x in tags.split(',') if x.strip()],dependencies=[],checklist=[{'name':x.strip(),'done':False} for x in checklist.splitlines() if x.strip()],estimated=estimated,actual=0,notes=notes,attachments=[],activity=[{'when':datetime.now().isoformat(timespec='minutes'),'by':'Current User','note':'Task created'}])); st.success('Task saved.'); st.rerun()
    c1,c2,c3,c4=st.columns([2,1,1,1]); query=c1.text_input('Search tasks, projects, or assignees'); proj=c2.selectbox('Project',['All Projects']+[p['name'] for p in P]); stat=c3.selectbox('Status',['All Statuses','Not Started','In Progress','Completed','On Hold']); pri=c4.selectbox('Priority',['All Priorities','Low','Medium','High','Critical'])
    data=[t for t in T if query.lower() in (t['name']+' '+t['assignee']+' '+pname(t['project_id'])).lower() and (proj=='All Projects' or pname(t['project_id'])==proj) and (stat=='All Statuses' or t['status']==stat) and (pri=='All Priorities' or t['priority']==pri)]
    if st.session_state.get('project_filter'): data=[t for t in data if t['project_id']==st.session_state.project_filter]; st.session_state.project_filter=None
    cols=st.columns(4)
    for c,(l,v) in zip(cols,[('Total Tasks',len(data)),('Completed',sum(t['status']=='Completed' for t in data)),('In Progress',sum(t['status']=='In Progress' for t in data)),('Overdue',sum(t['end']<today and t['status']!='Completed' for t in data))]):
        with c: metric(l,v)
    task_table(sorted(data,key=lambda t:(t['priority'] not in ('Critical','High'),t['end'])),'tasklist')
    if st.session_state.task_action:
        mode,tid=st.session_state.task_action; t=next((x for x in T if x['id']==tid),None)
        if t:
            with st.expander(('View Task' if mode=='view' else 'Edit Task')+' — '+t['name'],expanded=True):
                if mode=='view':
                    st.markdown(f"### {t['name']} · `{t['id']}`"); st.caption(f"{pname(t['project_id'])} · {t['status']} · {t['priority']} priority"); c1,c2,c3,c4=st.columns(4); c1.metric('Start',str(t['start'])); c2.metric('End',str(t['end'])); c3.metric('Duration',f"{t['duration']} days"); c4.metric('Complete',f"{t['percent']}%"); st.write(t['description']); st.write('Assignee:',t['assignee']); st.write('Tags:',', '.join(t['tags']) or 'None'); st.write('Progress notes:',t['notes'] or 'None'); st.write('Checklist / subtasks'); [st.checkbox(x['name'],value=x['done'],key=f"v_{tid}_{i}",disabled=True) for i,x in enumerate(t['checklist'])]
                    if st.button('Edit this task',key='edit_this'): st.session_state.task_action=('edit',tid); st.rerun()
                else:
                    with st.form('edit_task_form'):
                        name=st.text_input('Task name',t['name']); desc=st.text_area('Description',t['description']); c1,c2,c3=st.columns(3); start=c1.date_input('Start date',t['start']); end=c2.date_input('End date',t['end']); assignee=c3.selectbox('Assignee',[r['name'] for r in R],index=next((i for i,r in enumerate(R) if r['name']==t['assignee']),0)); c1,c2,c3=st.columns(3); status=c1.selectbox('Status',['Not Started','In Progress','On Hold','Completed'],index=['Not Started','In Progress','On Hold','Completed'].index(t['status']) if t['status'] in ['Not Started','In Progress','On Hold','Completed'] else 0); priority=c2.selectbox('Priority',['Low','Medium','High','Critical'],index=['Low','Medium','High','Critical'].index(t['priority'])); pct=c3.slider('% Complete',0,100,t['percent']); tags=st.text_input('Tags',', '.join(t['tags'])); notes=st.text_area('Progress notes',t['notes']); save=st.form_submit_button('Save Changes',type='primary')
                    if save:
                        if end<start: st.error('End date must not precede start date.')
                        else: t.update(name=name,description=desc,start=start,end=end,duration=(end-start).days+1,assignee=assignee,status=status,priority=priority,percent=pct,tags=[x.strip() for x in tags.split(',') if x.strip()],notes=notes); t['activity'].append({'when':datetime.now().isoformat(timespec='minutes'),'by':'Current User','note':'Task updated'}); st.session_state.task_action=None; st.success('Task updated.'); st.rerun()
                if st.button('Close details',key='close_details'): st.session_state.task_action=None; st.rerun()
elif page=='Gantt Chart':
    header('Gantt Chart','Visualize task dates, durations, and progress.')
    proj=st.selectbox('Project',['All Projects']+[p['name'] for p in P]); data=[t for t in T if proj=='All Projects' or pname(t['project_id'])==proj]
    if data:
        first=min(t['start'] for t in data); last=max(t['end'] for t in data); st.caption(f'Timeline: {first:%b %d, %Y} — {last:%b %d, %Y}')
        for t in sorted(data,key=lambda x:x['start']):
            st.markdown(f"**{t['id']} · {t['name']}** &nbsp; <span style='color:#71809a'>{t['assignee']} · {t['status']}</span>",unsafe_allow_html=True); st.progress(t['percent']/100,text=f"{t['start']:%b %d} – {t['end']:%b %d} · {t['percent']}%")
    st.caption('This starter timeline shows dates and progress; advanced critical-path scheduling and drag-to-reschedule are not included.')
elif page=='Resources':
    header('Resources','Manage team members, skills, availability, and workload.')
    with st.expander('＋ Add Resource'):
        with st.form('resource_form'):
            c1,c2=st.columns(2); name=c1.text_input('Full name *'); email=c2.text_input('Email address *'); c1,c2,c3=st.columns(3); role=c1.text_input('Role'); allocation=c2.number_input('Allocation (%)',0,200,0); status=c3.selectbox('Status',['Available','Allocated','On Leave']); skills=st.text_input('Skills, comma separated'); assigned=st.multiselect('Assign to projects',[p['name'] for p in P]); hours=st.number_input('Working hours per day',1,24,8); notes=st.text_area('Notes')
            if st.form_submit_button('Add Resource',type='primary'):
                if not name.strip() or not email.strip(): st.error('Name and email are required.')
                elif any(r['email'].lower()==email.lower() for r in R): st.error('Email already exists.')
                else: R.append(dict(id=uid('RES'),name=name,email=email,role=role,skills=[x.strip() for x in skills.split(',') if x.strip()],allocation=allocation,status='Overallocated' if allocation>100 else status,projects=[p['id'] for p in P if p['name'] in assigned],hours=hours,capacity=100,notes=notes)); st.success('Resource added.'); st.rerun()
    cols=st.columns(5)
    for c,(l,v) in zip(cols,[('Total Resources',len(R)),('Available',sum(r['status']=='Available' for r in R)),('Allocated',sum(r['status']=='Allocated' for r in R)),('Overallocated',sum(r['status']=='Overallocated' or r['allocation']>100 for r in R)),('On Leave',sum(r['status']=='On Leave' for r in R))]):
        with c: metric(l,v)
    a,b=st.columns(2)
    with a: st.subheader('Resource Allocation'); st.bar_chart(pd.DataFrame([{'Resource':r['name'],'Allocation':r['allocation']} for r in R]).set_index('Resource'))
    with b: st.subheader('Resource Status'); st.bar_chart(pd.Series({s:sum(r['status']==s for r in R) for s in ['Available','Allocated','Overallocated','On Leave']}))
    q=st.text_input('Search by name, role, skill, or project'); rolef=st.selectbox('Role filter',['All Roles']+sorted(set(r['role'] for r in R))); rows=[]
    for r in R:
        searchable=' '.join([r['name'],r['role'],r['email'],' '.join(r['skills']),' '.join(pname(pid) for pid in r['projects'])])
        if q.lower() in searchable.lower() and (rolef=='All Roles' or r['role']==rolef): rows.append({'Name':r['name'],'Email':r['email'],'Role':r['role'],'Skills':', '.join(r['skills']),'Current Project(s)':', '.join(pname(pid) for pid in r['projects']),'Allocation':f"{r['allocation']}%",'Availability':f"{max(0,100-r['allocation'])}%",'Status':r['status']})
    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    opts={r['name']+' — '+r['role']:r['id'] for r in R}; selected=st.selectbox('Select a resource to edit or delete',['']+list(opts))
    if selected:
        r=next(x for x in R if x['id']==opts[selected])
        with st.expander('Resource actions — '+r['name'],expanded=True):
            with st.form('edit_resource'):
                c1,c2=st.columns(2); name=c1.text_input('Name',r['name']); email=c2.text_input('Email',r['email']); c1,c2,c3=st.columns(3); role=c1.text_input('Role',r['role']); allocation=c2.number_input('Allocation (%)',0,200,r['allocation']); status=c3.selectbox('Status',['Available','Allocated','Overallocated','On Leave'],index=['Available','Allocated','Overallocated','On Leave'].index(r['status'])); skills=st.text_input('Skills',', '.join(r['skills'])); assigned=st.multiselect('Projects',[p['name'] for p in P],default=[pname(pid) for pid in r['projects']]); notes=st.text_area('Notes',r['notes']); save=st.form_submit_button('Save Resource',type='primary')
            if save: r.update(name=name,email=email,role=role,allocation=allocation,status='Overallocated' if allocation>100 else status,skills=[x.strip() for x in skills.split(',') if x.strip()],projects=[p['id'] for p in P if p['name'] in assigned],notes=notes); st.success('Resource updated.'); st.rerun()
            if st.button('Delete Resource',key='delete_'+r['id']): R.remove(r); st.rerun()
elif page=='Calendar':
    header('Calendar','View project timelines, task schedules, team availability, and events.')
    with st.expander('＋ Add Event'):
        with st.form('event_form'):
            c1,c2=st.columns(2); title=c1.text_input('Event title *'); typ=c2.selectbox('Event type',['Project Assignment','Meeting / Event','Leave','Milestone']); c1,c2=st.columns(2); project=c1.selectbox('Project',['None']+[p['name'] for p in P]); members=c2.multiselect('Resource(s)',[r['name'] for r in R]); c1,c2=st.columns(2); start=c1.date_input('Start date',today); end=c2.date_input('End date',today); desc=st.text_area('Description')
            if st.form_submit_button('Create Event',type='primary'):
                if not title.strip() or end<start: st.error('Enter a title and valid date range.')
                else: E.append(dict(id=uid('EVT'),title=title,type=typ,project=project,resources=members,start=start,end=end,description=desc)); st.success('Event created.'); st.rerun()
    month=st.date_input('Choose month',today.replace(day=1)); st.subheader(calendar.month_name[month.month]+' '+str(month.year)); weeks=calendar.Calendar(firstweekday=6).monthdatescalendar(month.year,month.month)
    for week in weeks:
        cols=st.columns(7)
        for col,day in zip(cols,week):
            with col:
                st.markdown(f"**{day.day}{'' if day.month==month.month else ' ·'}**")
                for t in [t for t in T if t['start']<=day<=t['end']][:2]: st.caption(('🟢 ' if t['status']=='Completed' else '🔵 ')+t['name'][:20])
                for ev in [e for e in E if e['start']<=day<=e['end']]: st.caption('🟣 '+ev['title'][:20])
                st.markdown('<hr style="margin:2px 0;border-color:#dce7f0">',unsafe_allow_html=True)
    st.subheader('Upcoming Events')
    if not E: st.info('No custom events yet.')
    for ev in sorted(E,key=lambda x:x['start']): st.write(f"**{ev['title']}** · {ev['type']} · {ev['start']} – {ev['end']} · {ev['project']}")
elif page=='Reports':
    header('Reports','Portfolio, schedule, task, and resource reporting.')
    a,b=st.columns(2)
    with a: st.subheader('Task Status'); st.bar_chart(pd.Series({s:sum(t['status']==s for t in T) for s in ['Not Started','In Progress','Completed','On Hold']}))
    with b: st.subheader('Resource Allocation'); st.bar_chart(pd.DataFrame([{'Resource':r['name'],'Allocation':r['allocation']} for r in R]).set_index('Resource'))
    report=[]
    for p in P:
        ts=[t for t in T if t['project_id']==p['id']]; report.append({'Project':p['name'],'Owner':p['owner'],'Status':p['status'],'Tasks':len(ts),'Completed':sum(t['status']=='Completed' for t in ts),'Average Complete':round(sum(t['percent'] for t in ts)/max(1,len(ts))),'Budget':p['budget']})
    st.dataframe(pd.DataFrame(report),use_container_width=True,hide_index=True); st.download_button('Download project report CSV',pd.DataFrame(report).to_csv(index=False),'vektor_project_report.csv','text/csv'); st.download_button('Download tasks CSV',pd.DataFrame([{'ID':t['id'],'Task':t['name'],'Project':pname(t['project_id']),'Status':t['status'],'Percent':t['percent'],'Assignee':t['assignee']} for t in T]).to_csv(index=False),'vektor_tasks.csv','text/csv')
elif page=='Settings':
    header('Settings','Workspace preferences and data management.'); st.text_input('Workspace name',value='Vektor Project Workspace'); st.selectbox('Default task view',['Task List','Gantt Chart','Calendar']); st.selectbox('Date format',['MMM DD, YYYY','DD/MM/YYYY','YYYY-MM-DD']); st.warning('Prototype data is stored in Streamlit session state. It resets after the server restarts and is not shared persistently across users.')
    payload={'projects':[{**p,'start':p['start'].isoformat(),'end':p['end'].isoformat()} for p in P],'resources':R,'tasks':[{**t,'start':t['start'].isoformat(),'end':t['end'].isoformat()} for t in T],'events':[{**e,'start':e['start'].isoformat(),'end':e['end'].isoformat()} for e in E]}; st.download_button('Export workspace JSON',json.dumps(payload,indent=2,default=str),'vektor_workspace.json','application/json')
    if st.button('Reset demo data'): st.session_state.projects,st.session_state.resources,st.session_state.tasks,st.session_state.events=seed(); st.rerun()
st.markdown('---'); st.caption('Vektor Project Manager · Streamlit prototype based on the supplied interface references. Add persistent storage, authentication, audit logging, and advanced scheduling before production use.')
