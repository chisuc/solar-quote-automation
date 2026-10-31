import os, json, uuid, secrets, hmac, mimetypes, asyncio
from pathlib import Path
from datetime import datetime, timezone
from urllib.parse import quote
import requests
from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.middleware.sessions import SessionMiddleware

USER=os.getenv("ADMIN_USER","admin")
PASS=os.getenv("ADMIN_PASSWORD","")
SESSION_SECRET=os.getenv("SESSION_SECRET",secrets.token_urlsafe(48))
SUPABASE_URL=os.getenv("SUPABASE_URL","").rstrip("/")
SUPABASE_SECRET_KEY=os.getenv("SUPABASE_SECRET_KEY","")
BUCKET=os.getenv("SUPABASE_BUCKET","customer-files")
app=FastAPI(title="Solar Quote Automation")
app.add_middleware(SessionMiddleware,secret_key=SESSION_SECRET,same_site="lax",https_only=os.getenv("COOKIE_SECURE","true")=="true")

# Reuse one headless Chromium process across GSA requests; pages/contexts remain request-scoped.
_gsa_playwright=None
_gsa_browser=None
_gsa_browser_lock=asyncio.Lock()
_gsa_capture_lock=asyncio.Lock()

def auth(r):
    if r.session.get("user")!=USER:
        raise HTTPException(401,"Vui lòng đăng nhập.")

def sb_headers(extra=None):
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        raise HTTPException(503,"Máy chủ chưa được cấu hình Supabase.")
    h={"apikey":SUPABASE_SECRET_KEY,"Authorization":f"Bearer {SUPABASE_SECRET_KEY}"}
    if extra: h.update(extra)
    return h

def sb_request(method, path, **kwargs):
    try:
        r=requests.request(method, SUPABASE_URL+path, headers=sb_headers(kwargs.pop("headers",{})), timeout=45, **kwargs)
    except requests.RequestException:
        raise HTTPException(502,"Không kết nối được Supabase.")
    if r.status_code >= 400:
        detail="Lỗi dịch vụ lưu trữ dữ liệu."
        try:
            j=r.json(); detail=j.get("message") or j.get("error") or j.get("msg") or detail
        except Exception: pass
        raise HTTPException(r.status_code, detail[:500])
    return r

def db(path): return "/rest/v1/"+path

def storage_path(pid, fid, filename):
    ext=Path(filename).suffix.lower()
    return f"{pid}/{fid}{ext}"

LOGIN="""<!doctype html><html lang='vi'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Đăng nhập | Solar Quote Automation</title><style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;font:15px Arial,sans-serif;color:#17382c;background:linear-gradient(135deg,#e9f3ee 0%,#f5f8f6 48%,#dcebe3 100%)}
.login-shell{width:min(440px,94vw)}.brand{text-align:center;margin-bottom:18px}.brand-mark{width:58px;height:58px;margin:0 auto 12px;border-radius:18px;background:#15392e;color:#fff;display:grid;place-items:center;font-size:29px;box-shadow:0 9px 24px #15392e26}.brand h1{font-size:22px;margin:0 0 7px;letter-spacing:.2px}.brand p{margin:0;color:#71827a;font-size:13px}
.card{background:#fff;padding:30px;border:1px solid #dce8e1;border-radius:18px;width:100%;box-shadow:0 18px 55px #17382c14}.card h2{font-size:19px;margin:0 0 7px}.sub{color:#71827a;font-size:13px;margin:0 0 24px}.field{margin:15px 0}.field label{display:block;font-size:13px;font-weight:700;margin-bottom:7px}.field input{width:100%;height:46px;padding:0 13px;border:1px solid #ccd9d2;border-radius:9px;font-size:15px;outline:none}.field input:focus{border-color:#16845b;box-shadow:0 0 0 3px #16845b19}.password-wrap{position:relative}.password-wrap input{padding-right:54px}.toggle{position:absolute;right:7px;top:7px;height:32px;width:42px;border:0;border-radius:7px;background:#eef5f1;color:#315b49;cursor:pointer;font-size:17px}.submit{width:100%;height:47px;margin-top:12px;border:0;border-radius:9px;background:#16845b;color:#fff;font-size:14px;font-weight:700;cursor:pointer;transition:.18s}.submit:hover{background:#116e4b;transform:translateY(-1px)}.submit:disabled{opacity:.65;cursor:wait;transform:none}.message{display:none;margin-top:14px;padding:11px 12px;border-radius:8px;font-size:13px;line-height:1.45}.message.error{display:block;color:#a52e2e;background:#fff0ef;border:1px solid #f2d2cf}.message.info{display:block;color:#245441;background:#edf7f1;border:1px solid #d2e9db}.foot{text-align:center;color:#89978f;font-size:11px;margin:18px 0 0}
</style></head><body><div class='login-shell'><div class='brand'><div class='brand-mark'>☀</div><h1>SOLAR QUOTE AUTOMATION</h1><p>Trung tâm quản lý & báo giá điện mặt trời</p></div><form class='card' id='loginForm'><h2>Đăng nhập hệ thống</h2><p class='sub'>Vui lòng nhập thông tin tài khoản để tiếp tục.</p><div class='field'><label for='username'>Tài khoản</label><input id='username' name='username' autocomplete='username' placeholder='Nhập tên tài khoản' required></div><div class='field'><label for='password'>Mật khẩu</label><div class='password-wrap'><input id='password' name='password' type='password' autocomplete='current-password' placeholder='Nhập mật khẩu' required><button class='toggle' type='button' id='togglePassword' aria-label='Hiện mật khẩu' title='Hiện mật khẩu'>◉</button></div></div><button class='submit' id='loginBtn' type='submit'>ĐĂNG NHẬP</button><div class='message' id='loginMessage' role='alert' aria-live='polite'></div></form><p class='foot'>BẢO MẬT · SOLAR QUOTE AUTOMATION</p></div><script>
const form=document.getElementById('loginForm'),pw=document.getElementById('password'),toggle=document.getElementById('togglePassword'),btn=document.getElementById('loginBtn'),msg=document.getElementById('loginMessage');
toggle.addEventListener('click',()=>{const show=pw.type==='password';pw.type=show?'text':'password';toggle.textContent=show?'⊘':'◉';toggle.title=show?'Ẩn mật khẩu':'Hiện mật khẩu';toggle.setAttribute('aria-label',toggle.title)});
form.addEventListener('submit',async e=>{e.preventDefault();msg.className='message';msg.textContent='';btn.disabled=true;btn.textContent='ĐANG KIỂM TRA...';
try{const r=await fetch('/login',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded','Accept':'application/json'},body:new URLSearchParams(new FormData(form))});let data={};try{data=await r.json()}catch(_){}
if(r.ok&&data.ok){msg.className='message info';msg.textContent='Đăng nhập thành công. Đang chuyển trang...';location.href='/';return}
msg.className='message error';msg.textContent=data.detail||'Không thể đăng nhập. Vui lòng thử lại.';
}catch(err){msg.className='message error';msg.textContent='Không kết nối được máy chủ. Kiểm tra Internet hoặc thử lại sau.'}
finally{btn.disabled=false;btn.textContent='ĐĂNG NHẬP'}});
</script></body></html>"""
PAGE="""<!doctype html><html lang='vi'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Solar Quote Automation</title><style>*{box-sizing:border-box}body{margin:0;background:#f3f6f5;font:14px Arial;color:#20372e}header{background:#15352b;color:white;padding:18px 24px}main{width:min(96vw,1580px);max-width:none;margin:24px auto;padding:0 14px}.panel{background:white;border:1px solid #dce5e1;border-radius:12px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.field label{display:block;font-size:12px;font-weight:bold;margin-bottom:6px}input,select{width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px}.btn{padding:10px 14px;background:#16845b;color:white;border:0;border-radius:7px;font-weight:bold;cursor:pointer}.alt{background:#e8f1ed;color:#245441}.row{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.project{padding:12px;border:1px solid #dce5e1;border-radius:8px;margin:8px 0;display:flex;justify-content:space-between}.muted{color:#687a73;font-size:12px}.gsa-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}.gsa-card{border:1px solid #dce5e1;border-radius:9px;padding:12px;min-width:0}.gsa-card h4{margin:0 0 12px;color:#245441}.gsa-metrics{width:100%;border-collapse:collapse}.gsa-metrics td{padding:7px 5px;border-bottom:1px solid #e8efeb}.gsa-metrics td:last-child{text-align:right;font-weight:bold}.gsa-sun svg{width:100%;height:auto}.gsa-card a{display:inline-block;margin-top:8px;color:#0879b9}@media(max-width:850px){.compass-field{grid-column:1/-1}.pv-chart-box{grid-column:1/-1;max-width:100%;min-height:250px}.compass-wrap{flex-wrap:wrap;justify-content:flex-start;gap:18px}.compass{width:210px!important;height:210px!important;min-width:210px;min-height:210px;flex-basis:210px}}@media(max-width:700px){.gsa-grid{grid-template-columns:1fr}}.compass-field{grid-column:1/3;min-width:0}.compass{width:220px!important;height:220px!important;min-width:220px;min-height:220px;max-width:none!important;max-height:none!important;aspect-ratio:1/1;flex:0 0 220px;border-radius:50%;box-sizing:border-box}.pv-chart-box{grid-column:3;grid-row:auto;align-self:center;display:flex;flex-direction:column;justify-content:center;min-width:0;width:100%;min-height:300px;margin:0!important;padding:14px 12px;overflow:hidden;border:1px solid #dce5e1;border-radius:10px;background:#fff}.pv-chart-box h4{font-size:13px;line-height:1.3;margin:0 0 8px;color:#245441}.pv-chart-box #pvChartCanvas{width:100%;overflow:hidden;flex:1;display:flex;align-items:center;justify-content:center}.pv-chart-box svg{display:block;width:100%;height:auto;min-width:0!important}.compass-wrap{display:flex;align-items:center;justify-content:center;gap:26px;flex-wrap:nowrap;padding:8px 4px}.compass{width:220px;height:220px;flex:0 0 220px}.compass-info{max-width:285px}.needle{height:68px}.needle i{border-bottom-width:60px}.needle b{top:60px}.compass-info{min-width:150px;flex:1}.compass-info .row{gap:6px}.compass-info .btn{font-size:12px;padding:8px 10px}.compass{width:250px;height:250px;max-width:75vw;max-height:75vw;border:3px solid #9aafa5;border-radius:50%;position:relative;background:radial-gradient(circle,#fff 0 57%,#edf4f0 58% 100%);box-shadow:inset 0 0 0 7px #f7faf8,0 2px 8px #00000012;touch-action:none;user-select:none;cursor:crosshair}.compass:before,.compass:after{content:'';position:absolute;background:#d7e2dc;left:50%;top:10%;width:1px;height:80%;transform:translateX(-50%)}.compass:after{transform:translateX(-50%) rotate(90deg)}.dir{position:absolute;font-size:11px;font-weight:bold;color:#52665d;z-index:2}.north{top:7px;left:50%;transform:translateX(-50%);color:#c0392b}.south{bottom:7px;left:50%;transform:translateX(-50%)}.east{right:8px;top:50%;transform:translateY(-50%)}.west{left:8px;top:50%;transform:translateY(-50%)}.ne{right:34px;top:32px}.se{right:34px;bottom:32px}.sw{left:34px;bottom:32px}.nw{left:34px;top:32px}.needle{position:absolute;z-index:3;left:50%;top:50%;width:8px;height:88px;transform:translate(-50%,-100%) rotate(180deg);transform-origin:50% 100%;pointer-events:none}.needle i{position:absolute;left:0;top:0;width:0;height:0;border-left:4px solid transparent;border-right:4px solid transparent;border-bottom:78px solid #e04435;filter:drop-shadow(0 1px 1px #777)}.needle b{position:absolute;left:1px;top:78px;width:6px;height:10px;background:#667b71;border-radius:0 0 3px 3px}.compass-center{position:absolute;z-index:4;left:50%;top:50%;width:15px;height:15px;border-radius:50%;background:#fff;border:4px solid #315c49;transform:translate(-50%,-50%);pointer-events:none}.compass-info{min-width:230px;flex:1}.compass-info .row{margin:12px 0}@media(max-width:650px){.grid{grid-template-columns:1fr 1fr}}@media(max-width:450px){.grid{grid-template-columns:1fr}}</style></head><body><header><b>☀ SOLAR QUOTE AUTOMATION</b> <button class='btn alt' onclick='logout()' style='float:right'>Đăng xuất</button></header><main><h1>Trung tâm báo giá điện mặt trời</h1><p class='muted'>Website MVP · Dữ liệu lưu trên máy chủ</p>
<div class='panel'><h3>Khách hàng & công trình</h3><div class='grid'><div class='field'><label>Tên dự án</label><input id='projectName'></div><div class='field'><label>Khách hàng / công ty</label><input id='customer'></div><div class='field'><label>Người liên hệ</label><input id='contact'></div><div class='field'><label>Số điện thoại</label><input id='phone'></div><div class='field'><label>Địa chỉ</label><input id='address'></div><div class='field'><label>Ngày báo giá</label><input id='date' type='date'></div><div class='field'><label>Diện tích mái (m²)</label><input id='roof' type='number'></div><div class='field'><label>Kết cấu mái</label><select id='roofType'><option>Mái tôn</option><option>Mái bê tông</option><option>Mái ngói</option></select></div></div></div>
<div class='panel'><h3>Cấu hình hệ thống</h3><div class='grid'><div class='field'><label>Thương hiệu tấm pin</label><input id='panelBrand'></div><div class='field'><label>Công suất tấm pin (Wp)</label><input id='panelW' type='number' value='625'></div><div class='field'><label>Số lượng tấm</label><input id='panelQty' type='number' value='30'></div><div class='field'><label>Công suất PV (kWp)</label><input id='pv' readonly></div><div class='field'><label>Inverter / Model</label><input id='invModel'></div><div class='field'><label>Công suất inverter (kW)</label><input id='invKw' type='number'></div><div class='field'><label>Số pha</label><select id='phase'><option>1 pha</option><option>3 pha</option></select></div><div class='field'><label>Pin lưu trữ / Model</label><input id='batModel'></div><div class='field'><label>Dung lượng mỗi bộ (kWh)</label><input id='batKwh' type='number' step='.01'></div><div class='field'><label>Số lượng bộ pin</label><input id='batQty' type='number' value='0'></div><div class='field'><label>Tổng dung lượng pin (kWh)</label><input id='batTotal' readonly></div></div></div>
<div class='panel'><h3>Điện năng & báo giá</h3><div class='grid'><div class='field'><label>Tiêu thụ tháng (kWh)</label><input id='load' type='number'></div><div class='field'><label>Sản lượng ĐMT năm (kWh)</label><input id='annual' type='number'></div><div class='field'><label>Tự tiêu thụ dự kiến (%)</label><input id='self' type='number' value='80'></div><div class='field'><label>Đơn giá điện (đ/kWh)</label><input id='price' type='number'></div><div class='field'><label>Tổng trước VAT (VNĐ)</label><input id='pre' type='number'></div><div class='field'><label>VAT (%)</label><input id='vat' type='number' value='8'></div><div class='field'><label>Tiền VAT</label><input id='vatMoney' readonly></div><div class='field'><label>Tổng sau VAT</label><input id='grand' readonly></div></div></div>
<div class='panel'><h3>Mô phỏng sản lượng PVWatts V8</h3><p class='muted'>Nhập tọa độ và thông số mô phỏng, sau đó bấm lấy dữ liệu trực tiếp từ PVWatts.</p><div class='grid'>
<div class='field' style='grid-column:1/-1'><label>Tọa độ công trình (dán Latitude, Longitude)</label><input id='coordPaste' type='text' placeholder='Ví dụ: 10.821006371676384, 106.68872440336703' autocomplete='off'><div id='coordStatus' class='muted'>Dán một dòng tọa độ Google Maps; website tự tách vĩ độ và kinh độ.</div><input id='lat' type='hidden'><input id='lon' type='hidden'></div>
<div class='field'><label>DC System Size (kW)</label><input id='pvSystemSize' type='number' step='.01' placeholder='Tự lấy từ công suất PV nếu để trống'></div>
<div class='field'><label>Module Type</label><select id='moduleType'><option value='0'>Standard</option><option value='1'>Premium</option><option value='2'>Thin film</option></select></div>
<div class='field'><label>Array Type</label><select id='arrayType'><option value='0' selected>Fixed - Open Rack</option><option value='1'>Fixed - Roof Mounted</option><option value='2'>1-Axis</option><option value='3'>1-Axis Backtracking</option><option value='4'>2-Axis</option></select></div>
<div class='field'><label>Tilt (độ)</label><input id='tilt' type='number' value='20' step='.1'></div>
<div class='field compass-field'><label>Hướng mặt tấm pin (Azimuth)</label><div class='compass-wrap'><div class='compass' id='compass' aria-label='La bàn chọn hướng PV'><span class='dir north'>BẮC</span><span class='dir east'>ĐÔNG</span><span class='dir south'>NAM</span><span class='dir west'>TÂY</span><span class='dir ne'>ĐB</span><span class='dir se'>ĐN</span><span class='dir sw'>TN</span><span class='dir nw'>TB</span><div class='needle' id='compassNeedle'><i></i><b></b></div><div class='compass-center'></div></div><div class='compass-info'><div><b id='azimuthText'>180° · Hướng Nam</b><div class='muted'>Xoay kim theo hướng mặt tấm pin thực tế đang quay.</div></div><div class='row'><button type='button' class='btn' id='confirmAzimuth' onclick='confirmAzimuth()'>Xác nhận hướng</button><button type='button' class='btn alt' id='editAzimuth' onclick='editAzimuth()'>Chỉnh sửa</button></div><div class='field' id='manualAzimuthWrap' style='display:none;max-width:180px'><label>Nhập góc thủ công (0–359,9°)</label><input id='azimuth' type='number' value='180' min='0' max='359.9' step='.1'></div><div class='muted' id='azimuthStatus'>Hướng hiện tại đã xác nhận.</div></div></div></div>
<div id='pvChart' class='pv-chart-box' style='display:none'><h4>Biểu đồ sản lượng theo tháng</h4><div id='pvChartCanvas'></div></div>
<div class='field'><label>System Losses (%)</label><input id='losses' type='number' value='14.08' step='.01'></div>
<div class='field'><label>DC to AC Size Ratio</label><input id='dcAcRatio' type='number' value='1.2' step='.01'></div>
<div class='field'><label>Inverter Efficiency (%)</label><input id='invEff' type='number' value='96' step='.1'></div>
<div class='field'><label>Ground Coverage Ratio</label><input id='gcr' type='number' value='.4' step='.01'></div>
<div class='field'><label>Albedo</label><select id='albedoMode'><option value='weather'>Theo dữ liệu thời tiết (Weather file)</option><option value='manual'>Nhập thủ công</option></select></div>
<div class='field' id='albedoField' style='display:none'><label>Albedo (0–1)</label><input id='albedo' type='number' value='.2' step='.01'></div>
</div><div class='row'><button class='btn' onclick='runPVWatts()'>Tính toán PVWatts</button><span id='pvStatus' class='muted'></span></div><div id='pvResult' class='muted'>Chưa có kết quả mô phỏng.</div></div>
<div class='panel'><h3>Ảnh phân tích tài nguyên mặt trời – Global Solar Atlas</h3><p class='muted'>Nhập tọa độ ở phần PVWatts, lưu dự án, sau đó hệ thống tự mở Global Solar Atlas và chụp 4 khu vực dữ liệu thực tế.</p><div class='row'><button class='btn' id='gsaCaptureBtn' onclick='captureGlobalSolarAtlas()'>Tự động lấy 4 hình ảnh</button><span id='gsaStatus' class='muted'>Chưa chụp ảnh.</span></div><div id='gsaResult' style='display:none'><div class='gsa-grid'><section class='gsa-card'><h4>1. Thông số bức xạ & PV</h4><img id='gsaImage1' alt='Thông số bức xạ và PV' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section><section class='gsa-card'><h4>2. Bản đồ vị trí công trình</h4><img id='gsaImage2' alt='Bản đồ vị trí công trình' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section><section class='gsa-card'><h4>3. Horizon & Sunpath</h4><img id='gsaImage3' alt='Horizon và Sunpath' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section><section class='gsa-card'><h4>4. Bản đồ tiềm năng PVOUT</h4><img id='gsaImage4' alt='Bản đồ PVOUT' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section></div><p class='muted'>Ảnh chụp trực tiếp từ Global Solar Atlas. Nguồn: © World Bank Group, ESMAP và Solargis.</p></div></div>
<div class='panel'><h3>Hồ sơ đầu vào</h3><input id='files' type='file' multiple><div class='row'><button class='btn' onclick='upload()'>Tải file lên máy chủ</button></div><div id='fileList' class='muted'></div></div>
<div class='panel'><h3>Dự án đã lưu</h3><div class='row'><button class='btn' onclick='save()'>Lưu dự án</button><button class='btn alt' onclick='fresh()'>Dự án mới</button><button class='btn alt' onclick='exportJson()'>Xuất JSON</button></div><div id='projects'>Đang tải...</div><div id='status' class='muted'></div></div></main>
<script>
let pid=null,pvwattsResult=null,gsaImages=[];const ids=['lat','lon','pvSystemSize','moduleType','arrayType','tilt','azimuth','losses','dcAcRatio','invEff','gcr','albedoMode','albedo','projectName','customer','contact','phone','address','date','roof','roofType','panelBrand','panelW','panelQty','invModel','invKw','phase','batModel','batKwh','batQty','load','annual','self','price','pre','vat'];const $=x=>document.getElementById(x),n=x=>Number($(x).value)||0;
function calc(){let kw=n('panelW')*n('panelQty')/1000,v=n('pre')*n('vat')/100;$('pv').value=kw.toFixed(3);$('batTotal').value=(n('batKwh')*n('batQty')).toFixed(2);$('vatMoney').value=v.toFixed(0);$('grand').value=(n('pre')+v).toFixed(0)}
document.addEventListener('input',e=>{if(e.target.tagName==='INPUT')calc()});
function payload(){let o={};ids.forEach(k=>o[k]=$(k).value);o.pvwattsResult=pvwattsResult;o.globalSolarAtlasResult=globalSolarAtlasResult;o.gsaImages=gsaImages;return o}function fill(o){ids.forEach(k=>{if(o[k]!==undefined)$(k).value=o[k]});pvwattsResult=o.pvwattsResult||null;globalSolarAtlasResult=o.globalSolarAtlasResult||null;gsaImages=o.gsaImages||[];renderGSAImages();renderPV();calc()}
$('albedoMode').addEventListener('change',()=>{$('albedoField').style.display=$('albedoMode').value==='manual'?'block':'none'});
let compassAngle=180,azimuthConfirmed=true;
function normAngle(v){return ((v%360)+360)%360}
function compassLabel(v){let dirs=['Bắc','Đông Bắc','Đông','Đông Nam','Nam','Tây Nam','Tây','Tây Bắc'];return dirs[Math.round(normAngle(v)/45)%8]}
function paintCompass(){let a=normAngle(compassAngle);$('compassNeedle').style.transform='translate(-50%,-100%) rotate('+a+'deg)';$('azimuthText').textContent=a.toFixed(1).replace('.0','')+'° · Hướng '+compassLabel(a);$('azimuth').value=a.toFixed(1);$('azimuthStatus').textContent=azimuthConfirmed?'Hướng đã xác nhận.':'Đang chỉnh hướng — nhấn Xác nhận hướng để áp dụng.'}
function setCompassFromPointer(e){let r=$('compass').getBoundingClientRect(),dx=e.clientX-(r.left+r.width/2),dy=e.clientY-(r.top+r.height/2);compassAngle=normAngle(Math.atan2(dx,-dy)*180/Math.PI);azimuthConfirmed=false;paintCompass()}
$('compass').addEventListener('pointerdown',e=>{setCompassFromPointer(e);$('compass').setPointerCapture(e.pointerId)});
$('compass').addEventListener('pointermove',e=>{if(e.buttons)setCompassFromPointer(e)});
function confirmAzimuth(){azimuthConfirmed=true;paintCompass();$('azimuthStatus').textContent='Đã xác nhận: '+$('azimuth').value+'° · '+compassLabel(compassAngle)}
function editAzimuth(){azimuthConfirmed=false;$('manualAzimuthWrap').style.display='block';$('azimuthStatus').textContent='Xoay kim trên la bàn hoặc nhập góc thủ công, sau đó xác nhận.'}
$('azimuth').addEventListener('input',()=>{let v=Number($('azimuth').value);if(Number.isFinite(v)&&v>=0&&v<360){compassAngle=v;azimuthConfirmed=false;paintCompass()}});
function parseCoordinates(){let raw=$('coordPaste').value.trim();let parts=raw.split(/[;,\\s]+/).filter(Boolean);if(parts.length!==2){$('coordStatus').textContent='Vui lòng dán đúng định dạng: vĩ độ, kinh độ.';return false}let lat=Number(parts[0]),lon=Number(parts[1]);if(!Number.isFinite(lat)||!Number.isFinite(lon)||lat< -90||lat>90||lon< -180||lon>180){$('coordStatus').textContent='Tọa độ không hợp lệ. Vĩ độ từ -90 đến 90, kinh độ từ -180 đến 180.';return false}$('lat').value=lat;$('lon').value=lon;$('coordStatus').textContent='Đã nhận tọa độ · Vĩ độ: '+lat+' · Kinh độ: '+lon;return true}
$('coordPaste').addEventListener('change',parseCoordinates);
$('coordPaste').addEventListener('blur',()=>{if($('coordPaste').value.trim())parseCoordinates()});
paintCompass();
function drawPVChart(a){let vals=a.map(Number),max=Math.max(...vals,1),W=520,H=260,L=48,R=8,T=28,B=34,plotW=W-L-R,plotH=H-T-B,step=plotW/12,bw=Math.min(27,step*.68),fmt=v=>Math.round(v).toLocaleString('vi-VN');let grid='',bars='';for(let i=0;i<=4;i++){let val=max*i/4,y=T+plotH-(plotH*i/4);grid+=`<line x1="${L}" y1="${y}" x2="${W-R}" y2="${y}" stroke="#e4ece8" stroke-width="1"/><text x="${L-6}" y="${y+4}" text-anchor="end" fill="#687a73" font-size="10">${fmt(val)}</text>`};vals.forEach((v,i)=>{let h=v/max*plotH,x=L+i*step+(step-bw)/2,y=T+plotH-h;bars+=`<rect x="${x}" y="${y}" width="${bw}" height="${h}" rx="3" fill="#16845b"><title>T${i+1}: ${fmt(v)} kWh</title></rect><text x="${x+bw/2}" y="${Math.max(17,y-5)}" text-anchor="middle" fill="#315c49" font-size="9">${fmt(v)}</text><text x="${x+bw/2}" y="${H-11}" text-anchor="middle" fill="#52665d" font-size="10">T${i+1}</text>`});$('pvChartCanvas').innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Biểu đồ sản lượng PVWatts 12 tháng" style="display:block;width:100%;height:auto;font-family:Arial,sans-serif">${grid}${bars}</svg>`;$('pvChart').style.display='flex'}
let globalSolarAtlasResult=null;
function renderGSAImages(){if(!gsaImages||!gsaImages.length){$('gsaResult').style.display='none';return}$('gsaResult').style.display='block';for(let i=0;i<4;i++){let item=gsaImages[i];$('gsaImage'+(i+1)).src=item?('/api/files/'+item.id):''}}
async function captureGlobalSolarAtlas(){try{if(!parseCoordinates())throw Error('Vui lòng kiểm tra tọa độ.');if(!pid)throw Error('Trước tiên hãy lưu dự án bằng nút Lưu dự án.');let b=$('gsaCaptureBtn');b.disabled=true;b.textContent='Đang mở Atlas và chụp ảnh...';$('gsaStatus').textContent='Đang khởi chạy trình duyệt tự động. Có thể mất 1–2 phút.';let r=await api('/api/projects/'+pid+'/gsa-capture',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({lat:n('lat'),lon:n('lon')})});gsaImages=r.images||[];renderGSAImages();$('gsaStatus').textContent='Đã chụp và lưu '+gsaImages.length+' ảnh vào dự án.';await listFiles()}catch(e){$('gsaStatus').textContent='Chưa hoàn tất: '+e.message}finally{$('gsaCaptureBtn').disabled=false;$('gsaCaptureBtn').textContent='Tự động lấy 4 hình ảnh'}}
function renderPV(){if(!pvwattsResult){$('pvResult').textContent='Chưa có kết quả mô phỏng.';$('pvChart').style.display='none';return}let a=pvwattsResult.monthly||[];$('pvResult').innerHTML='<b>Sản lượng AC năm: '+Number(pvwattsResult.annual||0).toLocaleString('vi-VN',{maximumFractionDigits:0})+' kWh</b><br>'+a.map((v,i)=>['T1','T2','T3','T4','T5','T6','T7','T8','T9','T10','T11','T12'][i]+': '+Number(v).toLocaleString('vi-VN',{maximumFractionDigits:0})+' kWh').join(' · ')+'<br><small>Trạm thời tiết: '+esc(pvwattsResult.station||'Chưa có thông tin')+'</small>';drawPVChart(a)}
async function runPVWatts(){try{if(!parseCoordinates())throw Error('Vui lòng kiểm tra lại tọa độ công trình.');if(!azimuthConfirmed)throw Error('Bạn hãy nhấn Xác nhận hướng trên la bàn trước khi tính toán.');$('pvStatus').textContent='Đang kết nối PVWatts...';let body={lat:n('lat'),lon:n('lon'),system_capacity:n('pvSystemSize')||n('pv'),module_type:n('moduleType'),array_type:n('arrayType'),tilt:n('tilt'),azimuth:n('azimuth'),losses:n('losses'),dc_ac_ratio:n('dcAcRatio'),inv_eff:n('invEff'),gcr:n('gcr'),use_wf_albedo:$('albedoMode').value==='weather'?1:0};if($('albedoMode').value==='manual')body.albedo=n('albedo');let r=await api('/api/pvwatts',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});pvwattsResult=r;$('annual').value=Math.round(r.annual||0);$('pvStatus').textContent='Mô phỏng thành công · PVWatts V'+r.version;renderPV()}catch(e){$('pvStatus').textContent='Không lấy được dữ liệu: '+e.message}}
async function api(u,opt={}){let r=await fetch(u,opt);if(r.status===401){location='/login-page';throw Error('Cần đăng nhập lại')}if(!r.ok)throw Error((await r.json()).detail||'Có lỗi');return r.json()}
async function save(){try{let p=payload(),r=pid?await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)}):await api('/api/projects',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});pid=r.id;$('status').textContent='Đã lưu lên máy chủ.';list()}catch(e){$('status').textContent=e.message}}
async function list(){let a=await api('/api/projects');$('projects').innerHTML=a.map(x=>`<div class='project'><span><b>${esc(x.name)}</b><br><small>${esc(x.customer||'')}</small></span><span><button class='btn alt' onclick=\"openP('${x.id}')\">Mở</button> <button class='btn alt' onclick=\"delP('${x.id}')\">Xóa</button></span></div>`).join('')||'Chưa có dự án'}
function esc(s){return String(s||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
async function openP(id){let x=await api('/api/projects/'+id);pid=id;fill(x.payload);let files=await listFiles();gsaImages=files.filter(v=>String(v.name||'').startsWith('global-solar-atlas-')).sort((a,b)=>a.name.localeCompare(b.name)).map(v=>({id:v.id,name:v.name}));renderGSAImages()}
async function delP(id){if(confirm('Xóa dự án?')){await api('/api/projects/'+id,{method:'DELETE'});if(pid===id)fresh();list()}}
function fresh(){pid=null;ids.forEach(k=>$(k).value=$(k).defaultValue||'');$('panelW').value=625;$('panelQty').value=30;$('batQty').value=0;$('vat').value=8;calc()}
async function upload(){if(!pid){alert('Hãy lưu dự án trước');return}let fd=new FormData();for(let f of $('files').files)fd.append('files',f);let r=await fetch('/api/projects/'+pid+'/files',{method:'POST',body:fd});if(!r.ok){alert('Tải file lỗi');return}$('files').value='';listFiles()}
async function listFiles(){if(!pid)return;let a=await api('/api/projects/'+pid+'/files');$('fileList').innerHTML=a.map(x=>`<div><a href='/api/files/${x.id}'>${esc(x.name)}</a></div>`).join('')||'Chưa có file';return a}
function exportJson(){let b=new Blob([JSON.stringify(payload(),null,2)],{type:'application/json'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='du-an.json';a.click()}
async function logout(){await fetch('/logout',{method:'POST'});location='/login-page'}
$('date').value=new Date().toISOString().slice(0,10);calc();list().catch(e=>$('status').textContent=e.message);
</script></main></body></html>"""

@app.get("/health")
def health():
    return {"status":"ok","storage":"supabase"}

@app.get("/login-page",response_class=HTMLResponse)
def login_page(): return LOGIN

@app.post("/login")
async def login(request:Request,username:str=Form(...),password:str=Form(...)):
    wants_json="application/json" in request.headers.get("accept","")
    if not USER or not PASS:
        detail="Máy chủ chưa cấu hình tài khoản đăng nhập. Vui lòng liên hệ quản trị viên."
        if wants_json: raise HTTPException(503,detail)
        return HTMLResponse(LOGIN.replace("</form>","<div class='message error'>"+detail+"</div></form>"),status_code=503)
    if username!=USER:
        detail="Tên tài khoản không chính xác. Vui lòng kiểm tra lại."
        if wants_json: raise HTTPException(401,detail)
        return HTMLResponse(LOGIN.replace("</form>","<div class='message error'>"+detail+"</div></form>"),status_code=401)
    if not hmac.compare_digest(password,PASS):
        detail="Mật khẩu không chính xác. Vui lòng thử lại."
        if wants_json: raise HTTPException(401,detail)
        return HTMLResponse(LOGIN.replace("</form>","<div class='message error'>"+detail+"</div></form>"),status_code=401)
    request.session.clear(); request.session["user"]=USER
    if wants_json: return {"ok":True}
    return RedirectResponse("/",status_code=303)

@app.post("/logout")
async def logout(request:Request):
    request.session.clear(); return {"ok":True}

@app.get("/",response_class=HTMLResponse)
def home(request:Request):
    if request.session.get("user")!=USER: return RedirectResponse("/login-page",status_code=303)
    return PAGE

@app.post("/api/pvwatts")
async def pvwatts(request:Request):
    auth(request)
    api_key=os.getenv("NLR_API_KEY","").strip()
    if not api_key:
        raise HTTPException(503,"Máy chủ chưa cấu hình NLR_API_KEY trong Render Environment.")
    try:
        p=await request.json()
        required=["lat","lon","system_capacity","module_type","array_type","tilt","azimuth","losses","dc_ac_ratio","inv_eff","gcr"]
        for k in required:
            if k not in p or p[k] is None: raise HTTPException(400,f"Thiếu thông số {k}.")
        if not (-90 <= float(p["lat"]) <= 90) or not (-180 <= float(p["lon"]) <= 180):
            raise HTTPException(400,"Tọa độ latitude/longitude không hợp lệ.")
        if float(p["system_capacity"]) <= 0:
            raise HTTPException(400,"DC System Size phải lớn hơn 0.")
        params={
            "api_key":api_key,"format":"json","lat":p["lat"],"lon":p["lon"],
            "system_capacity":p["system_capacity"],"module_type":p["module_type"],
            "array_type":p["array_type"],"tilt":p["tilt"],"azimuth":p["azimuth"],
            "losses":p["losses"],"dc_ac_ratio":p["dc_ac_ratio"],
            "inv_eff":p["inv_eff"],"gcr":p["gcr"],"timeframe":"monthly"
        }
        if int(p.get("use_wf_albedo",1))==1: params["use_wf_albedo"]=1
        elif p.get("albedo") is not None: params["albedo"]=p["albedo"]
        resp=requests.get("https://developer.nlr.gov/api/pvwatts/v8.json",params=params,timeout=60)
        try: data=resp.json()
        except ValueError: raise HTTPException(502,"PVWatts trả về dữ liệu không hợp lệ.")
        if resp.status_code>=400 or data.get("errors"):
            detail="; ".join(data.get("errors") or [f"PVWatts HTTP {resp.status_code}"])
            raise HTTPException(422,detail[:600])
        out=data.get("outputs",{})
        monthly=out.get("ac_monthly",[])
        if len(monthly)!=12: raise HTTPException(502,"PVWatts không trả đủ 12 giá trị sản lượng tháng.")
        station=data.get("station_info",{})
        return {"version":data.get("version"),"monthly":monthly,"annual":out.get("ac_annual"),
                "solrad_monthly":out.get("solrad_monthly"),"solrad_annual":out.get("solrad_annual"),
                "station":station.get("city") or station.get("location"),
                "warnings":data.get("warnings",[]),"inputs":data.get("inputs",{})}
    except HTTPException: raise
    except requests.RequestException as exc:
        print(f"PVWatts connection error: {type(exc).__name__}: {exc}", flush=True)
        raise HTTPException(502,"Render không kết nối được máy chủ PVWatts (NLR). Đã ghi chi tiết lỗi vào Render Logs; vui lòng thử lại sau khi Deploy bản cập nhật.")

@app.post("/api/projects/{pid}/gsa-capture")
async def capture_global_solar_atlas(pid:str,request:Request):
    global _gsa_playwright,_gsa_browser
    auth(request)
    from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
    p=await request.json()
    try:
        lat=float(p.get("lat")); lon=float(p.get("lon"))
        if not (-90<=lat<=90 and -180<=lon<=180): raise HTTPException(400,"Tọa độ không hợp lệ.")
    except (TypeError,ValueError): raise HTTPException(400,"Vui lòng nhập tọa độ hợp lệ.")
    check=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id"}).json()
    if not check: raise HTTPException(404,"Hãy lưu dự án trước khi chụp ảnh.")
    url=f"https://globalsolaratlas.info/detail?c={lat},{lon},10&m=site&s={lat},{lon}"
    labels=["Bảng dữ liệu bức xạ (Map data)","Bản đồ vệ tinh","Horizon & Sunpath","Bản đồ tiềm năng PVOUT"]
    try:
        # Serialize captures because a shared browser is reused; browser startup happens only once.
        async with _gsa_capture_lock:
            async with _gsa_browser_lock:
                if _gsa_playwright is None:
                    _gsa_playwright=await async_playwright().start()
                if _gsa_browser is None or not _gsa_browser.is_connected():
                    executable=_gsa_playwright.chromium.executable_path
                    if not os.path.exists(executable):
                        import subprocess,sys
                        install=subprocess.run([sys.executable,"-m","playwright","install","chromium"],capture_output=True,text=True,timeout=240)
                        if install.returncode!=0: raise RuntimeError("Không cài được Chromium: "+(install.stderr or install.stdout)[-1000:])
                    _gsa_browser=await _gsa_playwright.chromium.launch(headless=True,args=["--no-sandbox","--disable-dev-shm-usage","--disable-gpu","--disable-extensions"])
                browser=_gsa_browser
            context=await browser.new_context(viewport={"width":1920,"height":1200},device_scale_factor=1)
            page=await context.new_page()
            try:
                # domcontentloaded returns much earlier than networkidle on map-heavy pages.
                await page.goto(url,wait_until="domcontentloaded",timeout=60000)
                # Wait for the actual four source panels, not an arbitrary 22-second delay.
                await page.wait_for_function("""() => {
                  const body=(document.body?.innerText||'').toLowerCase();
                  const maps=[...document.querySelectorAll('.leaflet-container')].filter(e=>{const r=e.getBoundingClientRect();return r.width>250&&r.height>130});
                  const hasMetrics=/pvout specific/.test(body)&&/\\bdni\\b/.test(body)&&/\\bghi\\b/.test(body);
                  const hasSun=/solar azimuth|solar elevation|horizon and sunpath|horizon & sunpath/i.test(body);
                  return hasMetrics&&maps.length>=2&&hasSun;
                }""",timeout=35000)
                # Allow only a short tile/render settling window; don't wait for unrelated analytics.
                try:
                    await page.wait_for_function("""() => {
                      const maps=[...document.querySelectorAll('.leaflet-container')];
                      return maps.length>=2&&maps.every(m=>{
                        const tiles=[...m.querySelectorAll('img.leaflet-tile')];
                        return tiles.length===0||tiles.some(t=>t.complete&&t.naturalWidth>0);
                      });
                    }""",timeout=5000)
                except PlaywrightTimeoutError: pass
                found=await page.evaluate("""() => {
                  const visible=e=>{const r=e.getBoundingClientRect(),s=getComputedStyle(e);return r.width>250&&r.height>130&&s.display!=='none'&&s.visibility!=='hidden'};
                  const mark=(e,k)=>{if(e)e.setAttribute('data-gsa-shot',k)};
                  const texts=[...document.querySelectorAll('body *')].filter(e=>e.children.length===0&&/Map data/i.test(e.textContent||''));
                  let table=null;
                  for(const t of texts){let p=t;for(let n=0;n<7&&p;p=p.parentElement,n++){if(visible(p)&&/PVOUT specific/i.test(p.innerText||'')&&/DNI/i.test(p.innerText||'')&&/GHI/i.test(p.innerText||'')){table=p;break}}if(table)break}
                  if(table)mark(table,'metrics');
                  const maps=[...document.querySelectorAll('.leaflet-container')].filter(visible);
                  if(maps[0])mark(maps[0],'satellite');
                  if(maps.length>1)mark(maps[maps.length-1],'pvout');
                  const allEls=[...document.querySelectorAll('body *')];
                  const titleNodes=allEls.filter(e=>{
                    const tx=((e.innerText||'')+' '+(e.textContent||'')+' '+(e.getAttribute('aria-label')||'')+' '+(e.getAttribute('title')||'')).trim();
                    return tx.length>0&&tx.length<180&&/solar azimuth|solar elevation|horizon|sunpath|sun path/i.test(tx);
                  }).sort((a,b)=>(((a.innerText||'')+(a.textContent||'')).length)-(((b.innerText||'')+(b.textContent||'')).length));
                  let chart=null;
                  for(const t of titleNodes){
                    let q=t;
                    for(let n=0;n<12&&q;q=q.parentElement,n++){
                      if(!visible(q))continue;
                      const r=q.getBoundingClientRect(),txt=(q.innerText||'').toLowerCase();
                      const hasPlot=!!q.querySelector('svg,canvas');
                      const words=['north','east','south','west'].filter(w=>txt.includes(w)).length;
                      if(r.width>300&&r.height>160&&r.width<innerWidth*.98&&r.height<innerHeight*1.8&&(hasPlot||words>=2)){chart=q;break}
                    }
                    if(chart)break;
                  }
                  if(!chart){
                    const plots=[...document.querySelectorAll('svg,canvas')].filter(visible);
                    const ranked=plots.map(e=>{
                      let q=e,best=null;
                      for(let n=0;n<8&&q;q=q.parentElement,n++){
                        const r=q.getBoundingClientRect(),txt=(q.innerText||'').toLowerCase();
                        if(r.width<300||r.height<160||r.width>innerWidth*.98||r.height>innerHeight*1.8)continue;
                        const words=['north','east','south','west'].filter(w=>txt.includes(w)).length;
                        const hours=(txt.match(/\\b(?:[6-9]|1[0-8])h\\b/g)||[]).length;
                        const score=words*4+hours*2+(/solar azimuth|solar elevation|horizon|sunpath/.test(txt)?12:0);
                        if(!best||score>best.score)best={e:q,score,area:r.width*r.height};
                      }return best;
                    }).filter(x=>x&&x.score>=4).sort((a,b)=>b.score-a.score||a.area-b.area);
                    if(ranked.length)chart=ranked[0].e;
                  }
                  if(chart)mark(chart,'sunpath');
                  return {metrics:!!table,maps:maps.length,sunpath:!!chart};
                }""")
                required=["metrics","satellite","sunpath","pvout"]
                status=await page.evaluate("""() => Object.fromEntries(['metrics','satellite','sunpath','pvout'].map(k=>[k,!!document.querySelector('[data-gsa-shot="'+k+'"]')]))""")
                if not all(status.values()):
                    raise HTTPException(502,"Chưa tìm thấy đủ 4 panel gốc trên Global Solar Atlas. Phát hiện: "+str(status)+".")
                # Sequential element screenshots avoid scroll races; original PNG resolution is preserved.
                images=[]
                for key in required:
                    target=page.locator(f'[data-gsa-shot="{key}"]').first
                    await target.scroll_into_view_if_needed(timeout=10000)
                    data=await target.screenshot(type="png",timeout=25000,animations="disabled")
                    if not data or len(data)<3000: raise RuntimeError(f"Ảnh panel {key} rỗng.")
                    images.append(data)
            finally:
                await context.close()
        # Supabase storage writes are blocking requests; run four independent uploads concurrently.
        async def save_one(i,data):
            def upload_sync():
                fid=str(uuid.uuid4()); filename=f"global-solar-atlas-{i+1}.png"; path=storage_path(pid,fid,filename)
                sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(path,safe="/"),data=data,headers={"Content-Type":"image/png","x-upsert":"true"})
                try: sb_request("POST",db("uploads"),json={"id":fid,"pid":pid,"name":filename,"path":path},headers={"Prefer":"return=minimal"})
                except Exception:
                    sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[path]}); raise
                return {"id":fid,"name":filename,"label":labels[i]}
            return await asyncio.to_thread(upload_sync)
        saved=await asyncio.gather(*(save_one(i,data) for i,data in enumerate(images)))
        return {"images":saved,"source_url":url,"count":len(saved),"detected_panels":status}
    except HTTPException: raise
    except Exception as exc:
        print(f"Global Solar Atlas capture error: {type(exc).__name__}: {exc}",flush=True)
        raise HTTPException(502,f"Global Solar Atlas capture failed: {type(exc).__name__}: {str(exc)[:350]}")
    
@app.get("/api/projects")
def projects(request:Request):
    auth(request)
    r=sb_request("GET",db("projects"),params={"select":"id,name,customer,updated","order":"updated.desc"})
    return r.json()

@app.post("/api/projects")
async def create(request:Request):
    auth(request); p=await request.json()
    pid=str(uuid.uuid4()); now=datetime.now(timezone.utc).isoformat()
    name=p.get("projectName") or p.get("customer") or "Dự án mới"
    row={"id":pid,"name":name,"customer":p.get("customer",""),"payload":p,"updated":now}
    sb_request("POST",db("projects"),json=row,headers={"Prefer":"return=minimal"})
    return {"id":pid}

@app.get("/api/projects/{pid}")
def get_project(pid:str,request:Request):
    auth(request)
    r=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id,payload"})
    rows=r.json()
    if not rows: raise HTTPException(404,"Không tìm thấy dự án")
    return {"id":rows[0]["id"],"payload":rows[0]["payload"]}

@app.put("/api/projects/{pid}")
async def update(pid:str,request:Request):
    auth(request); p=await request.json(); now=datetime.now(timezone.utc).isoformat()
    name=p.get("projectName") or p.get("customer") or "Dự án mới"
    r=sb_request("PATCH",db("projects"),params={"id":f"eq.{pid}"},json={"name":name,"customer":p.get("customer",""),"payload":p,"updated":now},headers={"Prefer":"return=representation"})
    if not r.json(): raise HTTPException(404,"Không tìm thấy dự án")
    return {"id":pid}

@app.delete("/api/projects/{pid}")
def delete(pid:str,request:Request):
    auth(request)
    r=sb_request("GET",db("uploads"),params={"pid":f"eq.{pid}","select":"id,path"})
    rows=r.json()
    if rows:
        paths=[x["path"] for x in rows]
        sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":paths})
    sb_request("DELETE",db("uploads"),params={"pid":f"eq.{pid}"})
    sb_request("DELETE",db("projects"),params={"id":f"eq.{pid}"})
    return {"ok":True}

@app.post("/api/projects/{pid}/files")
async def upload(pid:str,request:Request,files:list[UploadFile]=File(...)):
    auth(request)
    check=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id"}).json()
    if not check: raise HTTPException(404,"Không tìm thấy dự án")
    out=[]; allowed={".pdf",".xlsx",".xls",".csv",".jpg",".jpeg",".png",".doc",".docx"}
    for f in files:
        name=Path(f.filename or "file").name; ext=Path(name).suffix.lower()
        if ext not in allowed: raise HTTPException(400,"Định dạng file chưa hỗ trợ")
        data=await f.read()
        if len(data)>20*1024*1024: raise HTTPException(413,"Mỗi file tối đa 20 MB")
        fid=str(uuid.uuid4()); path=storage_path(pid,fid,name)
        content_type=f.content_type or mimetypes.guess_type(name)[0] or "application/octet-stream"
        sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(path,safe="/"),data=data,headers={"Content-Type":content_type,"x-upsert":"false"})
        try:
            sb_request("POST",db("uploads"),json={"id":fid,"pid":pid,"name":name,"path":path},headers={"Prefer":"return=minimal"})
        except Exception:
            sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[path]})
            raise
        out.append({"id":fid,"name":name})
    return out

@app.get("/api/projects/{pid}/files")
def files(pid:str,request:Request):
    auth(request)
    r=sb_request("GET",db("uploads"),params={"pid":f"eq.{pid}","select":"id,name","order":"name.asc"})
    return r.json()

@app.get("/api/files/{fid}")
def download(fid:str,request:Request):
    auth(request)
    r=sb_request("GET",db("uploads"),params={"id":f"eq.{fid}","select":"name,path"})
    rows=r.json()
    if not rows: raise HTTPException(404,"Không tìm thấy file")
    item=rows[0]
    obj=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(item["path"],safe="/"))
    media=mimetypes.guess_type(item["name"])[0] or "application/octet-stream"
    return Response(content=obj.content,media_type=media,headers={"Content-Disposition":f"attachment; filename*=UTF-8''{quote(item['name'])}"})
