import os, sqlite3, json, uuid, secrets, hmac
from pathlib import Path
from datetime import datetime, timezone
from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from starlette.middleware.sessions import SessionMiddleware

DATA=Path(os.getenv("DATA_DIR","./data")); UP=DATA/"uploads"; DATA.mkdir(parents=True,exist_ok=True); UP.mkdir(exist_ok=True)
DB=DATA/"solar.db"; USER=os.getenv("ADMIN_USER","admin"); PASS=os.getenv("ADMIN_PASSWORD",""); SECRET=os.getenv("SESSION_SECRET",secrets.token_urlsafe(48))
app=FastAPI(title="Solar Quote Automation")
app.add_middleware(SessionMiddleware,secret_key=SECRET,same_site="lax",https_only=os.getenv("COOKIE_SECURE","false")=="true")
def conn():
 c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c
with conn() as c:
 c.execute("CREATE TABLE IF NOT EXISTS projects(id TEXT PRIMARY KEY,name TEXT,customer TEXT,payload TEXT,updated TEXT)")
 c.execute("CREATE TABLE IF NOT EXISTS uploads(id TEXT PRIMARY KEY,pid TEXT,name TEXT,stored TEXT)")
def auth(r):
 if r.session.get("user")!=USER: raise HTTPException(401,"Vui lòng đăng nhập.")
LOGIN="""<!doctype html><html lang='vi'><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Solar Quote</title><style>body{font:15px Arial;background:#eef4f1;display:grid;place-items:center;min-height:95vh;color:#17382c}.card{background:white;padding:30px;border-radius:14px;width:min(380px,90vw)}input,button{box-sizing:border-box;width:100%;padding:12px;margin:8px 0;border:1px solid #ccd9d2;border-radius:8px}button{background:#16845b;color:white;border:0;font-weight:bold}</style><form class='card' method='post' action='/login'><h2>☀ SOLAR QUOTE AUTOMATION</h2><p>Đăng nhập hệ thống</p><input name='username' placeholder='Tài khoản' required><input name='password' type='password' placeholder='Mật khẩu' required><button>ĐĂNG NHẬP</button>{}</form></html>"""
PAGE="""<!doctype html><html lang='vi'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Solar Quote Automation</title><style>*{box-sizing:border-box}body{margin:0;background:#f3f6f5;font:14px Arial;color:#20372e}header{background:#15352b;color:white;padding:18px 24px}main{max-width:1150px;margin:24px auto;padding:0 14px}.panel{background:white;border:1px solid #dce5e1;border-radius:12px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.field label{display:block;font-size:12px;font-weight:bold;margin-bottom:6px}input,select{width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px}.btn{padding:10px 14px;background:#16845b;color:white;border:0;border-radius:7px;font-weight:bold;cursor:pointer}.alt{background:#e8f1ed;color:#245441}.row{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.project{padding:12px;border:1px solid #dce5e1;border-radius:8px;margin:8px 0;display:flex;justify-content:space-between}.muted{color:#687a73;font-size:12px}@media(max-width:650px){.grid{grid-template-columns:1fr 1fr}}@media(max-width:450px){.grid{grid-template-columns:1fr}}</style></head><body><header><b>☀ SOLAR QUOTE AUTOMATION</b> <button class='btn alt' onclick='logout()' style='float:right'>Đăng xuất</button></header><main><h1>Trung tâm báo giá điện mặt trời</h1><p class='muted'>Website MVP · Dữ liệu lưu trên máy chủ</p>
<div class='panel'><h3>Khách hàng & công trình</h3><div class='grid'><div class='field'><label>Tên dự án</label><input id='projectName'></div><div class='field'><label>Khách hàng / công ty</label><input id='customer'></div><div class='field'><label>Người liên hệ</label><input id='contact'></div><div class='field'><label>Số điện thoại</label><input id='phone'></div><div class='field'><label>Địa chỉ</label><input id='address'></div><div class='field'><label>Ngày báo giá</label><input id='date' type='date'></div><div class='field'><label>Diện tích mái (m²)</label><input id='roof' type='number'></div><div class='field'><label>Kết cấu mái</label><select id='roofType'><option>Mái tôn</option><option>Mái bê tông</option><option>Mái ngói</option></select></div></div></div>
<div class='panel'><h3>Cấu hình hệ thống</h3><div class='grid'><div class='field'><label>Thương hiệu tấm pin</label><input id='panelBrand'></div><div class='field'><label>Công suất tấm pin (Wp)</label><input id='panelW' type='number' value='625'></div><div class='field'><label>Số lượng tấm</label><input id='panelQty' type='number' value='30'></div><div class='field'><label>Công suất PV (kWp)</label><input id='pv' readonly></div><div class='field'><label>Inverter / Model</label><input id='invModel'></div><div class='field'><label>Công suất inverter (kW)</label><input id='invKw' type='number'></div><div class='field'><label>Số pha</label><select id='phase'><option>1 pha</option><option>3 pha</option></select></div><div class='field'><label>Pin lưu trữ / Model</label><input id='batModel'></div><div class='field'><label>Dung lượng mỗi bộ (kWh)</label><input id='batKwh' type='number' step='.01'></div><div class='field'><label>Số lượng bộ pin</label><input id='batQty' type='number' value='0'></div><div class='field'><label>Tổng dung lượng pin (kWh)</label><input id='batTotal' readonly></div></div></div>
<div class='panel'><h3>Điện năng & báo giá</h3><div class='grid'><div class='field'><label>Tiêu thụ tháng (kWh)</label><input id='load' type='number'></div><div class='field'><label>Sản lượng ĐMT năm (kWh)</label><input id='annual' type='number'></div><div class='field'><label>Tự tiêu thụ dự kiến (%)</label><input id='self' type='number' value='80'></div><div class='field'><label>Đơn giá điện (đ/kWh)</label><input id='price' type='number'></div><div class='field'><label>Tổng trước VAT (VNĐ)</label><input id='pre' type='number'></div><div class='field'><label>VAT (%)</label><input id='vat' type='number' value='8'></div><div class='field'><label>Tiền VAT</label><input id='vatMoney' readonly></div><div class='field'><label>Tổng sau VAT</label><input id='grand' readonly></div></div></div>
<div class='panel'><h3>Hồ sơ đầu vào</h3><input id='files' type='file' multiple><div class='row'><button class='btn' onclick='upload()'>Tải file lên máy chủ</button></div><div id='fileList' class='muted'></div></div>
<div class='panel'><h3>Dự án đã lưu</h3><div class='row'><button class='btn' onclick='save()'>Lưu dự án</button><button class='btn alt' onclick='fresh()'>Dự án mới</button><button class='btn alt' onclick='exportJson()'>Xuất JSON</button></div><div id='projects'>Đang tải...</div><div id='status' class='muted'></div></div></main>
<script>
let pid=null;const ids=['projectName','customer','contact','phone','address','date','roof','roofType','panelBrand','panelW','panelQty','invModel','invKw','phase','batModel','batKwh','batQty','load','annual','self','price','pre','vat'];const $=x=>document.getElementById(x),n=x=>Number($(x).value)||0;
function calc(){let kw=n('panelW')*n('panelQty')/1000,v=n('pre')*n('vat')/100;$('pv').value=kw.toFixed(3);$('batTotal').value=(n('batKwh')*n('batQty')).toFixed(2);$('vatMoney').value=v.toFixed(0);$('grand').value=(n('pre')+v).toFixed(0)}
document.addEventListener('input',e=>{if(e.target.tagName==='INPUT')calc()});
function payload(){let o={};ids.forEach(k=>o[k]=$(k).value);return o}function fill(o){ids.forEach(k=>{if(o[k]!==undefined)$(k).value=o[k]});calc()}
async function api(u,opt={}){let r=await fetch(u,opt);if(r.status===401){location='/login-page';throw Error('Cần đăng nhập lại')}if(!r.ok)throw Error((await r.json()).detail||'Có lỗi');return r.json()}
async function save(){try{let p=payload(),r=pid?await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)}):await api('/api/projects',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});pid=r.id;$('status').textContent='Đã lưu lên máy chủ.';list()}catch(e){$('status').textContent=e.message}}
async function list(){let a=await api('/api/projects');$('projects').innerHTML=a.map(x=>`<div class='project'><span><b>${esc(x.name)}</b><br><small>${esc(x.customer||'')}</small></span><span><button class='btn alt' onclick=\"openP('${x.id}')\">Mở</button> <button class='btn alt' onclick=\"delP('${x.id}')\">Xóa</button></span></div>`).join('')||'Chưa có dự án'}
function esc(s){return String(s||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
async function openP(id){let x=await api('/api/projects/'+id);pid=id;fill(x.payload);listFiles()}
async function delP(id){if(confirm('Xóa dự án?')){await api('/api/projects/'+id,{method:'DELETE'});if(pid===id)fresh();list()}}
function fresh(){pid=null;ids.forEach(k=>$(k).value=$(k).defaultValue||'');$('panelW').value=625;$('panelQty').value=30;$('batQty').value=0;$('vat').value=8;calc()}
async function upload(){if(!pid){alert('Hãy lưu dự án trước');return}let fd=new FormData();for(let f of $('files').files)fd.append('files',f);let r=await fetch('/api/projects/'+pid+'/files',{method:'POST',body:fd});if(!r.ok){alert('Tải file lỗi');return}$('files').value='';listFiles()}
async function listFiles(){if(!pid)return;let a=await api('/api/projects/'+pid+'/files');$('fileList').innerHTML=a.map(x=>`<div><a href='/api/files/${x.id}'>${esc(x.name)}</a></div>`).join('')||'Chưa có file'}
function exportJson(){let b=new Blob([JSON.stringify(payload(),null,2)],{type:'application/json'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='du-an.json';a.click()}
async function logout(){await fetch('/logout',{method:'POST'});location='/login-page'}
$('date').value=new Date().toISOString().slice(0,10);calc();list().catch(e=>$('status').textContent=e.message);
</script></main></body></html>"""
@app.get("/health")
def health(): return {"status":"ok"}
@app.get("/login-page",response_class=HTMLResponse)
def login_page(): return LOGIN.format("")
@app.post("/login")
async def login(request:Request,username:str=Form(...),password:str=Form(...)):
 if username==USER and PASS and hmac.compare_digest(password,PASS):
  request.session.clear(); request.session["user"]=USER; return RedirectResponse("/",status_code=303)
 return HTMLResponse(LOGIN.format("<p style='color:red'>Sai tài khoản hoặc mật khẩu</p>"),status_code=401)
@app.post("/logout")
async def logout(request:Request): request.session.clear(); return {"ok":True}
@app.get("/",response_class=HTMLResponse)
def home(request:Request):
 if request.session.get("user")!=USER: return RedirectResponse("/login-page",status_code=303)
 return PAGE
@app.get("/api/projects")
def projects(request:Request):
 auth(request)
 with conn() as c: rows=c.execute("SELECT id,name,customer,updated FROM projects ORDER BY updated DESC").fetchall()
 return [dict(x) for x in rows]
@app.post("/api/projects")
async def create(request:Request):
 auth(request); p=await request.json(); pid=str(uuid.uuid4()); now=datetime.now(timezone.utc).isoformat(); name=p.get("projectName") or p.get("customer") or "Dự án mới"
 with conn() as c: c.execute("INSERT INTO projects VALUES(?,?,?,?,?,?)",(pid,name,p.get("customer",""),json.dumps(p,ensure_ascii=False),now))
 return {"id":pid}
@app.get("/api/projects/{pid}")
def get_project(pid:str,request:Request):
 auth(request)
 with conn() as c: row=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone()
 if not row: raise HTTPException(404,"Không tìm thấy dự án")
 return {"id":row["id"],"payload":json.loads(row["payload"])}
@app.put("/api/projects/{pid}")
async def update(pid:str,request:Request):
 auth(request); p=await request.json(); now=datetime.now(timezone.utc).isoformat(); name=p.get("projectName") or p.get("customer") or "Dự án mới"
 with conn() as c:
  cur=c.execute("UPDATE projects SET name=?,customer=?,payload=?,updated=? WHERE id=?",(name,p.get("customer",""),json.dumps(p,ensure_ascii=False),now,pid))
  if cur.rowcount==0: raise HTTPException(404,"Không tìm thấy dự án")
 return {"id":pid}
@app.delete("/api/projects/{pid}")
def delete(pid:str,request:Request):
 auth(request)
 with conn() as c:
  rows=c.execute("SELECT stored FROM uploads WHERE pid=?",(pid,)).fetchall()
  for x in rows: (UP/x["stored"]).unlink(missing_ok=True)
  c.execute("DELETE FROM uploads WHERE pid=?",(pid,)); c.execute("DELETE FROM projects WHERE id=?",(pid,))
 return {"ok":True}
@app.post("/api/projects/{pid}/files")
async def upload(pid:str,request:Request,files:list[UploadFile]=File(...)):
 auth(request)
 with conn() as c:
  if not c.execute("SELECT 1 FROM projects WHERE id=?",(pid,)).fetchone(): raise HTTPException(404,"Không tìm thấy dự án")
 out=[]; allowed={".pdf",".xlsx",".xls",".csv",".jpg",".jpeg",".png",".doc",".docx"}
 for f in files:
  name=Path(f.filename or "file").name; ext=Path(name).suffix.lower()
  if ext not in allowed: raise HTTPException(400,"Định dạng file chưa hỗ trợ")
  data=await f.read()
  if len(data)>20*1024*1024: raise HTTPException(413,"Mỗi file tối đa 20 MB")
  stored=secrets.token_hex(16)+ext; (UP/stored).write_bytes(data); fid=str(uuid.uuid4())
  with conn() as c: c.execute("INSERT INTO uploads VALUES(?,?,?,?)",(fid,pid,name,stored))
  out.append({"id":fid,"name":name})
 return out
@app.get("/api/projects/{pid}/files")
def files(pid:str,request:Request):
 auth(request)
 with conn() as c: rows=c.execute("SELECT id,name FROM uploads WHERE pid=?",(pid,)).fetchall()
 return [dict(x) for x in rows]
@app.get("/api/files/{fid}")
def download(fid:str,request:Request):
 auth(request)
 with conn() as c: row=c.execute("SELECT name,stored FROM uploads WHERE id=?",(fid,)).fetchone()
 if not row: raise HTTPException(404,"Không tìm thấy file")
 path=UP/row["stored"]
 if not path.exists(): raise HTTPException(404,"File không còn")
 return FileResponse(path,filename=row["name"])
