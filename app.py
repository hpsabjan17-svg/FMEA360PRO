from flask import Flask, render_template, request, redirect, url_for, send_file, session, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from io import BytesIO
from datetime import date
import sqlite3, os, re

app=Flask(__name__)
app.secret_key=os.environ.get('SECRET_KEY','change-this-secret-in-render')
app.config.update(SESSION_COOKIE_HTTPONLY=True,SESSION_COOKIE_SAMESITE='Lax',SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE','false').lower()=='true')
DB=os.path.join('database','fmea.db')

def db():
    os.makedirs('database',exist_ok=True); c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def addcol(c,t,n,d):
    if n not in [x[1] for x in c.execute(f'PRAGMA table_info({t})')]: c.execute(f'ALTER TABLE {t} ADD COLUMN {n} {d}')

def setup():
    c=db(); q=c.cursor()
    q.execute('CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT UNIQUE,password_hash TEXT,created_date TEXT)')
    for n,d in [('name',"TEXT DEFAULT ''"),('organisation',"TEXT DEFAULT 'Antolin'"),('language',"TEXT DEFAULT 'English'")]: addcol(q,'users',n,d)
    q.execute('CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY AUTOINCREMENT,project_name TEXT,product_name TEXT,customer TEXT,project_number TEXT,created_date TEXT)')
    for n,d in [('oem_name',"TEXT DEFAULT 'Generic'"),('compliance_mode',"TEXT DEFAULT 'AIAG-VDA 2019'"),('user_id','INTEGER')]: addcol(q,'projects',n,d)
    q.execute('CREATE TABLE IF NOT EXISTS oem_standards(id INTEGER PRIMARY KEY AUTOINCREMENT,oem_name TEXT UNIQUE,standard_framework TEXT,cc_symbol TEXT,sc_symbol TEXT,archiving_period_years INTEGER,description TEXT)')
    for x in [('Volkswagen Group','AIAG-VDA','D/TLD','K',15),('BMW Group','AIAG-VDA','DS','PTC',12),('Ford Motor Co','AIAG-VDA','∇','SC',10),('General Motors','AIAG-VDA','KPC','PQC',10),('Stellantis','AIAG-VDA','S','R',10),('Generic','AIAG-VDA 2019','CC','SC',10)]: q.execute('INSERT OR IGNORE INTO oem_standards(oem_name,standard_framework,cc_symbol,sc_symbol,archiving_period_years) VALUES(?,?,?,?,?)',x)
    q.execute('CREATE TABLE IF NOT EXISTS functional_analysis(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,function TEXT,requirement TEXT)')
    for n,d in [('parent_id','INTEGER'),('level','INTEGER DEFAULT 0'),('display_number',"TEXT DEFAULT ''"),('node_type',"TEXT DEFAULT 'Function'")]: addcol(q,'functional_analysis',n,d)
    q.execute('CREATE TABLE IF NOT EXISTS boundary_diagram(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,external_element TEXT,interaction TEXT,direction TEXT,description TEXT)')
    q.execute('CREATE TABLE IF NOT EXISTS product_structure(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,parent_id INTEGER,component_name TEXT,component_type TEXT,label TEXT,part_number TEXT,level INTEGER DEFAULT 0,description TEXT)')
    for n,d in [('display_number',"TEXT DEFAULT ''"),('node_text',"TEXT DEFAULT ''")]: addcol(q,'product_structure',n,d)
    q.execute('CREATE TABLE IF NOT EXISTS key_characteristics(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,component_id INTEGER,characteristic TEXT,specification TEXT,tolerance TEXT,severity INTEGER,responsibility TEXT)')
    for n,d in [('sc_type',"TEXT DEFAULT 'SC'"),('point_label',"TEXT DEFAULT ''"),('line_group',"TEXT DEFAULT ''")]: addcol(q,'key_characteristics',n,d)
    q.execute('CREATE TABLE IF NOT EXISTS functional_links(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,function_id INTEGER,component_id INTEGER,requirement TEXT)')
    q.execute('CREATE TABLE IF NOT EXISTS dfmea(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,component_id INTEGER,function TEXT,failure_mode TEXT,failure_effect TEXT,severity INTEGER,cause TEXT,occurrence INTEGER,prevention_control TEXT,detection_control TEXT,detection INTEGER,rpn INTEGER,recommended_action TEXT,responsibility TEXT,target_date TEXT,action_status TEXT)')
    q.execute('CREATE TABLE IF NOT EXISTS pfmea(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,component_id INTEGER,process_step TEXT,process_function TEXT,failure_mode TEXT,failure_effect TEXT,severity INTEGER,cause TEXT,occurrence INTEGER,prevention_control TEXT,detection_control TEXT,detection INTEGER,rpn INTEGER,recommended_action TEXT,responsibility TEXT,target_date TEXT,action_status TEXT)')
    q.execute('CREATE TABLE IF NOT EXISTS control_plan(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER,component_id INTEGER,process_step TEXT,characteristic TEXT,specification TEXT,control_method TEXT,measurement_method TEXT,sample_size TEXT,frequency TEXT,responsibility TEXT,reaction_plan TEXT)')
    q.execute('UPDATE projects SET user_id=(SELECT MIN(id) FROM users) WHERE user_id IS NULL AND EXISTS(SELECT 1 FROM users)')
    c.commit(); c.close()

def uid(): return session.get('user_id')
def owned(c,p): return c.execute('SELECT * FROM projects WHERE id=? AND user_id=?',(p,uid())).fetchone()
def projects(c): return c.execute('SELECT * FROM projects WHERE user_id=? ORDER BY id DESC',(uid(),)).fetchall()
def goodpw(p): return len(p)>=8 and re.search(r'[A-Z]',p) and re.search(r'\d',p) and re.search(r'[^A-Za-z0-9]',p)

@app.before_request
def auth():
    if request.endpoint in {'login','register','forgot_password','static'}: return
    if not uid(): return redirect(url_for('register'))

@app.route('/register',methods=['GET','POST'])
def register():
    if uid(): return redirect(url_for('dashboard'))
    error=None
    if request.method=='POST':
        name=request.form.get('name','').strip(); org=request.form.get('organisation','').strip(); user=request.form.get('username','').strip(); p=request.form.get('password',''); cp=request.form.get('confirm_password','')
        if not name or not org or not user or not p: error='Please fill in all required fields.'
        elif org.lower()!='antolin': error='Account creation is allowed only for Antolin organisation.'
        elif p!=cp: error='Passwords do not match.'
        elif not goodpw(p): error='Password must be at least 8 characters and contain 1 uppercase letter, 1 number and 1 special symbol.'
        else:
            c=db()
            if c.execute('SELECT id FROM users WHERE username=?',(user,)).fetchone(): error='Username already exists.'
            else:
                lang=request.form.get('language','English'); c.execute('INSERT INTO users(name,organisation,language,username,password_hash,created_date) VALUES(?,?,?,?,?,?)',(name,'Antolin',lang,user,generate_password_hash(p),date.today().isoformat())); c.commit(); c.close(); return redirect(url_for('login',created=1))
            c.close()
    return render_template('register.html',error=error)

@app.route('/login',methods=['GET','POST'])
def login():
    if uid(): return redirect(url_for('dashboard'))
    error=None
    if request.method=='POST':
        c=db(); u=c.execute('SELECT * FROM users WHERE username=?',(request.form.get('username','').strip(),)).fetchone(); c.close()
        if u and check_password_hash(u['password_hash'],request.form.get('password','')):
            session.clear(); session.update(user_id=u['id'],username=u['username'],name=u['name'],organisation=u['organisation'],language=u['language']); return redirect(url_for('dashboard'))
        error='Invalid username or password.'
    return render_template('login.html',error=error,created=request.args.get('created')=='1')

@app.route('/forgot-password',methods=['GET','POST'])
def forgot_password():
    error=success=None
    if request.method=='POST':
        p=request.form.get('password',''); cp=request.form.get('confirm_password',''); c=db()
        u=c.execute('SELECT * FROM users WHERE username=? AND name=? AND organisation=?',(request.form.get('username','').strip(),request.form.get('name','').strip(),'Antolin')).fetchone()
        if not u: error='Account details could not be verified.'
        elif p!=cp: error='Passwords do not match.'
        elif not goodpw(p): error='Password must be at least 8 characters and contain 1 uppercase letter, 1 number and 1 special symbol.'
        else: c.execute('UPDATE users SET password_hash=? WHERE id=?',(generate_password_hash(p),u['id'])); c.commit(); success='Password changed successfully. You can now log in.'
        c.close()
    return render_template('forgot_password.html',error=error,success=success)

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/')
@app.route('/dashboard')
def dashboard():
    c=db(); ps=projects(c); c.close(); return render_template('dashboard.html',projects=ps)

@app.route('/project',methods=['GET','POST'])
def project():
    c=db()
    if request.method=='POST':
        c.execute('INSERT INTO projects(project_name,product_name,customer,oem_name,compliance_mode,project_number,created_date,user_id) VALUES(?,?,?,?,?,?,?,?)',(request.form.get('project_name','').strip(),request.form.get('product_name',''),request.form.get('customer','').strip(),request.form.get('oem_name','Generic'),request.form.get('compliance_mode','AIAG-VDA 2019'),request.form.get('project_number','').strip(),request.form.get('created_date') or date.today().isoformat(),uid())); c.commit(); c.close(); return redirect(url_for('project'))
    ps=projects(c); oems=c.execute('SELECT * FROM oem_standards ORDER BY oem_name').fetchall(); c.close(); return render_template('project.html',projects=ps,oems=oems,today=date.today().isoformat())

def module_data(c,pid):
    ps=projects(c); comps=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; return ps,comps

@app.route('/product-structure',methods=['GET','POST'])
def product_structure():
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id',''); parent=request.form.get('parent_id') or None; name=request.form.get('component_name','').strip()
        if pid and owned(c,pid) and name:
            level=0
            if parent:
                r=c.execute('SELECT level FROM product_structure WHERE id=? AND project_id=?',(parent,pid)).fetchone(); level=(r['level']+1) if r else 0
            if not request.form.get('display_number'): request.form['display_number']=''
            c.execute('INSERT INTO product_structure(project_id,parent_id,component_name,component_type,label,part_number,level,description,display_number,node_text) VALUES(?,?,?,?,?,?,?,?,?,?)',(pid,parent,name,request.form.get('component_type','Assembly'),request.form.get('label',''),request.form.get('part_number',''),level,request.form.get('description',''),request.form.get('display_number'),name)); c.commit()
        c.close(); return redirect(url_for('product_structure',project_id=pid))
    ps=projects(c); nodes=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; c.close(); return render_template('product_structure.html',projects=ps,nodes=nodes,selected_project_id=pid)

@app.route('/functional-analysis',methods=['GET','POST'])
def functional_analysis():
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id',''); parent=request.form.get('parent_id') or None; fun=request.form.get('function','').strip()
        if pid and owned(c,pid) and fun:
            level=0
            if parent:
                r=c.execute('SELECT level FROM functional_analysis WHERE id=? AND project_id=?',(parent,pid)).fetchone(); level=(r['level']+1) if r else 0
            c.execute('INSERT INTO functional_analysis(project_id,function,requirement,parent_id,level,display_number,node_type) VALUES(?,?,?,?,?,?,?)',(pid,fun,request.form.get('requirement',''),parent,level,request.form.get('display_number',''),request.form.get('node_type','Function'))); c.commit()
        c.close(); return redirect(url_for('functional_analysis',project_id=pid))
    ps=projects(c); nodes=c.execute('SELECT * FROM functional_analysis WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; c.close(); return render_template('functional_analysis.html',projects=ps,nodes=nodes,selected_project_id=pid)

@app.post('/api/node/<table>/<int:rid>')
def api_node(table,rid):
    if table not in {'product_structure','functional_analysis'}: return jsonify(ok=False),400
    c=db(); row=c.execute(f'SELECT t.id FROM {table} t JOIN projects p ON t.project_id=p.id WHERE t.id=? AND p.user_id=?',(rid,uid())).fetchone()
    if not row: c.close(); return jsonify(ok=False),403
    d=request.get_json() or {}; allowed={'display_number','component_name','function','requirement'}; parts=[]; vals=[]
    for k in allowed:
        if k in d: parts.append(f'{k}=?'); vals.append(str(d[k]))
    if parts: vals.append(rid); c.execute(f'UPDATE {table} SET {",".join(parts)} WHERE id=?',vals); c.commit()
    c.close(); return jsonify(ok=True)

@app.route('/key-characteristics',methods=['GET','POST'])
def key_characteristics():
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id',''); cid=request.form.get('component_id',''); ch=request.form.get('characteristic','').strip()
        if pid and cid and ch and owned(c,pid):
            c.execute('INSERT INTO key_characteristics(project_id,component_id,characteristic,specification,tolerance,severity,responsibility,sc_type,point_label,line_group) VALUES(?,?,?,?,?,?,?,?,?,?)',(pid,cid,ch,request.form.get('specification',''),request.form.get('tolerance',''),int(request.form.get('severity','1') or 1),request.form.get('responsibility',''),request.form.get('sc_type','SC'),request.form.get('point_label',''),request.form.get('line_group',''))); c.commit()
        c.close(); return redirect(url_for('key_characteristics',project_id=pid))
    ps=projects(c); comps=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; rec=c.execute('SELECT k.*,ps.component_name FROM key_characteristics k LEFT JOIN product_structure ps ON k.component_id=ps.id JOIN projects p ON k.project_id=p.id WHERE p.user_id=? ORDER BY k.id DESC',(uid(),)).fetchall(); c.close(); return render_template('key_characteristics.html',projects=ps,components=comps,records=rec,selected_project_id=pid)

@app.post('/key-characteristics/delete/<int:rid>')
def kc_delete(rid):
    c=db(); c.execute('DELETE FROM key_characteristics WHERE id IN (SELECT k.id FROM key_characteristics k JOIN projects p ON k.project_id=p.id WHERE k.id=? AND p.user_id=?)',(rid,uid())); c.commit(); c.close(); return jsonify(ok=True)

@app.route('/functional-links',methods=['GET','POST'])
def functional_links():
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id',''); f=request.form.get('function_id'); p=request.form.get('component_id')
        if pid and f and p and owned(c,pid):
            old=c.execute('SELECT id FROM functional_links WHERE project_id=? AND function_id=? AND component_id=?',(pid,f,p)).fetchone()
            if old: c.execute('DELETE FROM functional_links WHERE id=?',(old['id'],))
            else: c.execute('INSERT INTO functional_links(project_id,function_id,component_id,requirement) VALUES(?,?,?,?)',(pid,f,p,''))
            c.commit()
        c.close(); return redirect(url_for('functional_links',project_id=pid))
    ps=projects(c); fs=c.execute('SELECT * FROM functional_analysis WHERE project_id=? ORDER BY id',(pid,)).fetchall() if pid else []; products=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; links=c.execute('SELECT function_id,component_id FROM functional_links WHERE project_id=?',(pid,)).fetchall() if pid else []; linked={(x['function_id'],x['component_id']) for x in links}; c.close(); return render_template('functional_links.html',projects=ps,functions=fs,products=products,linked=linked,selected_project_id=pid)

@app.route('/dfmea',methods=['GET','POST'])
def dfmea():
    return fmea_page('dfmea')
@app.route('/pfmea',methods=['GET','POST'])
def pfmea():
    return fmea_page('pfmea')

def fmea_page(kind):
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id','')
        if pid and owned(c,pid) and request.form.get('failure_mode','').strip():
            try: s=max(1,min(10,int(request.form.get('severity','1')))); o=max(1,min(10,int(request.form.get('occurrence','1')))); d=max(1,min(10,int(request.form.get('detection','1'))))
            except ValueError: s=o=d=1
            if kind=='dfmea':
                c.execute('INSERT INTO dfmea(project_id,component_id,function,failure_mode,failure_effect,severity,cause,occurrence,prevention_control,detection_control,detection,rpn,recommended_action,responsibility,target_date,action_status) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(pid,request.form.get('component_id') or None,request.form.get('function',''),request.form.get('failure_mode'),request.form.get('failure_effect'),s,request.form.get('cause'),o,request.form.get('prevention_control'),request.form.get('detection_control'),d,s*o*d,request.form.get('recommended_action'),request.form.get('responsibility'),request.form.get('target_date'),request.form.get('action_status','Open')))
            else:
                c.execute('INSERT INTO pfmea(project_id,component_id,process_step,process_function,failure_mode,failure_effect,severity,cause,occurrence,prevention_control,detection_control,detection,rpn,recommended_action,responsibility,target_date,action_status) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(pid,request.form.get('component_id') or None,request.form.get('process_step'),request.form.get('process_function'),request.form.get('failure_mode'),request.form.get('failure_effect'),s,request.form.get('cause'),o,request.form.get('prevention_control'),request.form.get('detection_control'),d,s*o*d,request.form.get('recommended_action'),request.form.get('responsibility'),request.form.get('target_date'),request.form.get('action_status','Open')))
            c.commit()
        c.close(); return redirect(url_for(kind,project_id=pid))
    ps=projects(c); comps=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []
    table='dfmea' if kind=='dfmea' else 'pfmea'; rec=c.execute(f'SELECT x.*,ps.component_name FROM {table} x LEFT JOIN product_structure ps ON x.component_id=ps.id JOIN projects p ON x.project_id=p.id WHERE p.user_id=? ORDER BY x.id DESC',(uid(),)).fetchall(); c.close(); return render_template(kind+'.html',projects=ps,components=comps,records=rec,selected_project_id=pid)

@app.route('/control-plan',methods=['GET','POST'])
def control_plan():
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id','')
        if pid and owned(c,pid) and request.form.get('characteristic','').strip():
            keys=['component_id','process_step','characteristic','specification','control_method','measurement_method','sample_size','frequency','responsibility','reaction_plan']; vals=[request.form.get(k,'') for k in keys]; c.execute('INSERT INTO control_plan(project_id,component_id,process_step,characteristic,specification,control_method,measurement_method,sample_size,frequency,responsibility,reaction_plan) VALUES(?,?,?,?,?,?,?,?,?,?,?)',(pid,*vals)); c.commit()
        c.close(); return redirect(url_for('control_plan',project_id=pid))
    ps=projects(c); comps=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; rec=c.execute('SELECT cp.*,ps.component_name FROM control_plan cp LEFT JOIN product_structure ps ON cp.component_id=ps.id JOIN projects p ON cp.project_id=p.id WHERE p.user_id=? ORDER BY cp.id DESC',(uid(),)).fetchall(); c.close(); return render_template('control_plan.html',projects=ps,components=comps,records=rec,selected_project_id=pid)

@app.route('/boundary-diagram',methods=['GET','POST'])
def boundary_diagram():
    c=db(); pid=request.args.get('project_id','')
    if pid and not owned(c,pid): pid=''
    if request.method=='POST':
        pid=request.form.get('project_id','')
        if pid and owned(c,pid) and request.form.get('external_element','').strip(): c.execute('INSERT INTO boundary_diagram(project_id,external_element,interaction,direction,description) VALUES(?,?,?,?,?)',(pid,request.form.get('external_element'),request.form.get('interaction'),request.form.get('direction'),request.form.get('description'))); c.commit()
        c.close(); return redirect(url_for('boundary_diagram',project_id=pid))
    ps=projects(c); rec=c.execute('SELECT * FROM boundary_diagram WHERE project_id=? ORDER BY id DESC',(pid,)).fetchall() if pid else []; c.close(); return render_template('boundary_diagram.html',projects=ps,boundaries=rec,selected_project_id=pid)

@app.route('/reports')
def reports():
    c=db(); pid=request.args.get('project_id',''); info=owned(c,pid) if pid else None; ps=projects(c); summary={}; tables=['functional_analysis','boundary_diagram','product_structure','key_characteristics','functional_links','dfmea','pfmea','control_plan']
    for t in tables: summary[t]=c.execute(f'SELECT COUNT(*) FROM {t} WHERE project_id=?',(pid,)).fetchone()[0] if pid else 0
    comps=c.execute('SELECT * FROM product_structure WHERE project_id=? ORDER BY level,id',(pid,)).fetchall() if pid else []; df=c.execute('SELECT d.*,ps.component_name FROM dfmea d LEFT JOIN product_structure ps ON d.component_id=ps.id WHERE d.project_id=?',(pid,)).fetchall() if pid else []; pf=c.execute('SELECT p.*,ps.component_name FROM pfmea p LEFT JOIN product_structure ps ON p.component_id=ps.id WHERE p.project_id=?',(pid,)).fetchall() if pid else []; c.close(); return render_template('reports.html',projects=ps,project_info=info,summary=summary,components=comps,dfmea_records=df,pfmea_records=pf,selected_project_id=pid)

@app.route('/export-excel')
def export_excel():
    pid=request.args.get('project_id',''); c=db(); p=owned(c,pid)
    if not p: c.close(); return 'Project not found',404
    wb=Workbook(); ws=wb.active; ws.title='Project'; ws.append(['Field','Value'])
    for k in ['project_name','product_name','customer','oem_name','compliance_mode','project_number','created_date']: ws.append([k.replace('_',' ').title(),p[k]])
    sets=[('Product Structure',['Level','Number','Component','Type','Part Number','Description'],'SELECT level,display_number,component_name,component_type,part_number,description FROM product_structure WHERE project_id=?'),('Functional Analysis',['Level','Number','Function','Requirement'],'SELECT level,display_number,function,requirement FROM functional_analysis WHERE project_id=?'),('DFMEA',['Component','Function','Failure Mode','Effect','S','Cause','O','D','RPN','Action'],'SELECT ps.component_name,d.function,d.failure_mode,d.failure_effect,d.severity,d.cause,d.occurrence,d.detection,d.rpn,d.recommended_action FROM dfmea d LEFT JOIN product_structure ps ON d.component_id=ps.id WHERE d.project_id=?'),('PFMEA',['Component','Process Step','Function','Failure Mode','Effect','S','O','D','RPN','Action'],'SELECT ps.component_name,p.process_step,p.process_function,p.failure_mode,p.failure_effect,p.severity,p.occurrence,p.detection,p.rpn,p.recommended_action FROM pfmea p LEFT JOIN product_structure ps ON p.component_id=ps.id WHERE p.project_id=?')]
    for name,heads,sql in sets:
        s=wb.create_sheet(name); s.append(heads)
        for r in c.execute(sql,(pid,)).fetchall(): s.append(list(r))
    for s in wb.worksheets:
        for cell in s[1]: cell.font=__import__('openpyxl').styles.Font(bold=True,color='FFFFFF'); cell.fill=__import__('openpyxl').styles.PatternFill('solid',fgColor='17365D')
        s.freeze_panes='A2'
    c.close(); out=BytesIO(); wb.save(out); out.seek(0); return send_file(out,as_attachment=True,download_name='Automotive_FMEA_Report.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

setup()
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=False)
