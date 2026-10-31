import os, json, uuid, secrets, hmac, mimetypes
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

LOGIN="""<!doctype html><html lang='vi'><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Solar Quote</title><style>body{font:15px Arial;background:#eef4f1;display:grid;place-items:center;min-height:95vh;color:#17382c}.card{background:white;padding:30px;border-radius:14px;width:min(380px,90vw)}input,button{box-sizing:border-box;width:100%;padding:12px;margin:8px 0;border:1px solid #ccd9d2;border-radius:8px}button{background:#16845b;color:white;border:0;font-weight:bold}</style><form class='card' method='post' action='/login'><h2>☀ SOLAR QUOTE AUTOMATION</h2><p>Đăng nhập hệ thống</p><input name='username' placeholder='Tài khoản' required><input name='password' type='password' placeholder='Mật khẩu' required><button>ĐĂNG NHẬP</button>{}</form></html>"""
PAGE="""<!doctype html><html lang='vi'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Solar Quote Automation</title><style>*{box-sizing:border-box}body{margin:0;background:#f3f6f5;font:14px Arial;color:#20372e}header{background:#15352b;color:white;padding:18px 24px}main{max-width:1150px;margin:24px auto;padding:0 14px}.panel{background:white;border:1px solid #dce5e1;border-radius:12px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.field label{display:block;font-size:12px;font-weight:bold;margin-bottom:6px}input,select{width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px}.btn{padding:10px 14px;background:#16845b;color:white;border:0;border-radius:7px;font-weight:bold;cursor:pointer}.alt{background:#e8f1ed;color:#245441}.row{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.project{padding:12px;border:1px solid #dce5e1;border-radius:8px;margin:8px 0;display:flex;justify-content:space-between}.muted{color:#687a73;font-size:12px}.gsa-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}.gsa-card{border:1px solid #dce5e1;border-radius:9px;padding:12px;min-width:0}.gsa-card h4{margin:0 0 12px;color:#245441}.gsa-metrics{width:100%;border-collapse:collapse}.gsa-metrics td{padding:7px 5px;border-bottom:1px solid #e8efeb}.gsa-metrics td:last-child{text-align:right;font-weight:bold}.gsa-sun svg{width:100%;height:auto}.gsa-card a{display:inline-block;margin-top:8px;color:#0879b9}@media(max-width:700px){.gsa-grid{grid-template-columns:1fr}}.compass-wrap{display:flex;align-items:center;gap:28px;flex-wrap:wrap;padding:12px 0}.compass{width:250px;height:250px;max-width:75vw;max-height:75vw;border:3px solid #9aafa5;border-radius:50%;position:relative;background:radial-gradient(circle,#fff 0 57%,#edf4f0 58% 100%);box-shadow:inset 0 0 0 7px #f7faf8,0 2px 8px #00000012;touch-action:none;user-select:none;cursor:crosshair}.compass:before,.compass:after{content:'';position:absolute;background:#d7e2dc;left:50%;top:10%;width:1px;height:80%;transform:translateX(-50%)}.compass:after{transform:translateX(-50%) rotate(90deg)}.dir{position:absolute;font-size:11px;font-weight:bold;color:#52665d;z-index:2}.north{top:7px;left:50%;transform:translateX(-50%);color:#c0392b}.south{bottom:7px;left:50%;transform:translateX(-50%)}.east{right:8px;top:50%;transform:translateY(-50%)}.west{left:8px;top:50%;transform:translateY(-50%)}.ne{right:34px;top:32px}.se{right:34px;bottom:32px}.sw{left:34px;bottom:32px}.nw{left:34px;top:32px}.needle{position:absolute;z-index:3;left:50%;top:50%;width:8px;height:88px;transform:translate(-50%,-100%) rotate(180deg);transform-origin:50% 100%;pointer-events:none}.needle i{position:absolute;left:0;top:0;width:0;height:0;border-left:4px solid transparent;border-right:4px solid transparent;border-bottom:78px solid #e04435;filter:drop-shadow(0 1px 1px #777)}.needle b{position:absolute;left:1px;top:78px;width:6px;height:10px;background:#667b71;border-radius:0 0 3px 3px}.compass-center{position:absolute;z-index:4;left:50%;top:50%;width:15px;height:15px;border-radius:50%;background:#fff;border:4px solid #315c49;transform:translate(-50%,-50%);pointer-events:none}.compass-info{min-width:230px;flex:1}.compass-info .row{margin:12px 0}@media(max-width:650px){.grid{grid-template-columns:1fr 1fr}}@media(max-width:450px){.grid{grid-template-columns:1fr}}</style></head><body><header><b>☀ SOLAR QUOTE AUTOMATION</b> <button class='btn alt' onclick='logout()' style='float:right'>Đăng xuất</button></header><main><h1>Trung tâm báo giá điện mặt trời</h1><p class='muted'>Website MVP · Dữ liệu lưu trên máy chủ</p>
<div class='panel'><h3>Khách hàng & công trình</h3><div class='grid'><div class='field'><label>Tên dự án</label><input id='projectName'></div><div class='field'><label>Khách hàng / công ty</label><input id='customer'></div><div class='field'><label>Người liên hệ</label><input id='contact'></div><div class='field'><label>Số điện thoại</label><input id='phone'></div><div class='field'><label>Địa chỉ</label><input id='address'></div><div class='field'><label>Ngày báo giá</label><input id='date' type='date'></div><div class='field'><label>Diện tích mái (m²)</label><input id='roof' type='number'></div><div class='field'><label>Kết cấu mái</label><select id='roofType'><option>Mái tôn</option><option>Mái bê tông</option><option>Mái ngói</option></select></div></div></div>
<div class='panel'><h3>Cấu hình hệ thống</h3><div class='grid'><div class='field'><label>Thương hiệu tấm pin</label><input id='panelBrand'></div><div class='field'><label>Công suất tấm pin (Wp)</label><input id='panelW' type='number' value='625'></div><div class='field'><label>Số lượng tấm</label><input id='panelQty' type='number' value='30'></div><div class='field'><label>Công suất PV (kWp)</label><input id='pv' readonly></div><div class='field'><label>Inverter / Model</label><input id='invModel'></div><div class='field'><label>Công suất inverter (kW)</label><input id='invKw' type='number'></div><div class='field'><label>Số pha</label><select id='phase'><option>1 pha</option><option>3 pha</option></select></div><div class='field'><label>Pin lưu trữ / Model</label><input id='batModel'></div><div class='field'><label>Dung lượng mỗi bộ (kWh)</label><input id='batKwh' type='number' step='.01'></div><div class='field'><label>Số lượng bộ pin</label><input id='batQty' type='number' value='0'></div><div class='field'><label>Tổng dung lượng pin (kWh)</label><input id='batTotal' readonly></div></div></div>
<div class='panel'><h3>Điện năng & báo giá</h3><div class='grid'><div class='field'><label>Tiêu thụ tháng (kWh)</label><input id='load' type='number'></div><div class='field'><label>Sản lượng ĐMT năm (kWh)</label><input id='annual' type='number'></div><div class='field'><label>Tự tiêu thụ dự kiến (%)</label><input id='self' type='number' value='80'></div><div class='field'><label>Đơn giá điện (đ/kWh)</label><input id='price' type='number'></div><div class='field'><label>Tổng trước VAT (VNĐ)</label><input id='pre' type='number'></div><div class='field'><label>VAT (%)</label><input id='vat' type='number' value='8'></div><div class='field'><label>Tiền VAT</label><input id='vatMoney' readonly></div><div class='field'><label>Tổng sau VAT</label><input id='grand' readonly></div></div></div>
<div class='panel'><h3>Mô phỏng sản lượng PVWatts V8</h3><p class='muted'>Nhập tọa độ và thông số mô phỏng, sau đó bấm lấy dữ liệu trực tiếp từ PVWatts.</p><div class='grid'>
<div class='field' style='grid-column:1/-1'><label>Tọa độ công trình (dán Latitude, Longitude)</label><input id='coordPaste' type='text' placeholder='Ví dụ: 10.821006371676384, 106.68872440336703' autocomplete='off'><div id='coordStatus' class='muted'>Dán một dòng tọa độ Google Maps; website tự tách vĩ độ và kinh độ.</div><input id='lat' type='hidden'><input id='lon' type='hidden'></div>
<div class='field'><label>DC System Size (kW)</label><input id='pvSystemSize' type='number' step='.01' placeholder='Tự lấy từ công suất PV nếu để trống'></div>
<div class='field'><label>Module Type</label><select id='moduleType'><option value='0'>Standard</option><option value='1'>Premium</option><option value='2'>Thin film</option></select></div>
<div class='field'><label>Array Type</label><select id='arrayType'><option value='0' selected>Fixed - Open Rack</option><option value='1'>Fixed - Roof Mounted</option><option value='2'>1-Axis</option><option value='3'>1-Axis Backtracking</option><option value='4'>2-Axis</option></select></div>
<div class='field'><label>Tilt (độ)</label><input id='tilt' type='number' value='20' step='.1'></div>
<div class='field' style='grid-column:1/-1'><label>Hướng mặt tấm pin (Azimuth)</label><div class='compass-wrap'><div class='compass' id='compass' aria-label='La bàn chọn hướng PV'><span class='dir north'>BẮC</span><span class='dir east'>ĐÔNG</span><span class='dir south'>NAM</span><span class='dir west'>TÂY</span><span class='dir ne'>ĐB</span><span class='dir se'>ĐN</span><span class='dir sw'>TN</span><span class='dir nw'>TB</span><div class='needle' id='compassNeedle'><i></i><b></b></div><div class='compass-center'></div></div><div class='compass-info'><div><b id='azimuthText'>180° · Hướng Nam</b><div class='muted'>Xoay kim theo hướng mặt tấm pin thực tế đang quay.</div></div><div class='row'><button type='button' class='btn' id='confirmAzimuth' onclick='confirmAzimuth()'>Xác nhận hướng</button><button type='button' class='btn alt' id='editAzimuth' onclick='editAzimuth()'>Chỉnh sửa</button></div><div class='field' id='manualAzimuthWrap' style='display:none;max-width:180px'><label>Nhập góc thủ công (0–359,9°)</label><input id='azimuth' type='number' value='180' min='0' max='359.9' step='.1'></div><div class='muted' id='azimuthStatus'>Hướng hiện tại đã xác nhận.</div></div></div></div>
<div class='field'><label>System Losses (%)</label><input id='losses' type='number' value='14.08' step='.01'></div>
<div class='field'><label>DC to AC Size Ratio</label><input id='dcAcRatio' type='number' value='1.2' step='.01'></div>
<div class='field'><label>Inverter Efficiency (%)</label><input id='invEff' type='number' value='96' step='.1'></div>
<div class='field'><label>Ground Coverage Ratio</label><input id='gcr' type='number' value='.4' step='.01'></div>
<div class='field'><label>Albedo</label><select id='albedoMode'><option value='weather'>Theo dữ liệu thời tiết (Weather file)</option><option value='manual'>Nhập thủ công</option></select></div>
<div class='field' id='albedoField' style='display:none'><label>Albedo (0–1)</label><input id='albedo' type='number' value='.2' step='.01'></div>
</div><div class='row'><button class='btn' onclick='runPVWatts()'>Tính toán PVWatts</button><span id='pvStatus' class='muted'></span></div><div id='pvResult' class='muted'>Chưa có kết quả mô phỏng.</div><div id='pvChart' style='display:none;margin-top:18px'><h4 style='margin:0 0 10px'>Biểu đồ sản lượng điện mặt trời theo tháng</h4><div id='pvChartCanvas' style='width:100%;overflow-x:auto'></div></div></div>
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
function drawPVChart(a){let vals=a.map(Number),max=Math.max(...vals,1),W=900,H=330,L=64,R=18,T=28,B=48,plotW=W-L-R,plotH=H-T-B,step=plotW/12,bw=Math.min(42,step*.62),fmt=v=>Math.round(v).toLocaleString('vi-VN');let grid='',bars='';for(let i=0;i<=4;i++){let val=max*i/4,y=T+plotH-(plotH*i/4);grid+=`<line x1="${L}" y1="${y}" x2="${W-R}" y2="${y}" stroke="#e4ece8" stroke-width="1"/><text x="${L-9}" y="${y+4}" text-anchor="end" fill="#687a73" font-size="11">${fmt(val)}</text>`};vals.forEach((v,i)=>{let h=v/max*plotH,x=L+i*step+(step-bw)/2,y=T+plotH-h;bars+=`<rect x="${x}" y="${y}" width="${bw}" height="${h}" rx="4" fill="#16845b"><title>T${i+1}: ${fmt(v)} kWh</title></rect><text x="${x+bw/2}" y="${Math.max(13,y-7)}" text-anchor="middle" fill="#315c49" font-size="10">${fmt(v)}</text><text x="${x+bw/2}" y="${H-19}" text-anchor="middle" fill="#52665d" font-size="12">T${i+1}</text>`});$('pvChartCanvas').innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Biểu đồ sản lượng PVWatts 12 tháng" style="display:block;width:100%;min-width:650px;height:auto;font-family:Arial,sans-serif">${grid}${bars}<text x="${L}" y="14" fill="#52665d" font-size="11">Sản lượng (kWh)</text></svg>`;$('pvChart').style.display='block'}
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
def login_page(): return LOGIN.replace("{}", "")

@app.post("/login")
async def login(request:Request,username:str=Form(...),password:str=Form(...)):
    if username==USER and PASS and hmac.compare_digest(password,PASS):
        request.session.clear(); request.session["user"]=USER
        return RedirectResponse("/",status_code=303)
    return HTMLResponse(LOGIN.replace("{}", "<p style='color:red'>Sai tài khoản hoặc mật khẩu</p>"),status_code=401)

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
    auth(request)
    import re, time
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
    p=await request.json()
    try:
        lat=float(p.get("lat")); lon=float(p.get("lon"))
        if not (-90<=lat<=90 and -180<=lon<=180):
            raise HTTPException(400,"Tọa độ không hợp lệ.")
    except (TypeError,ValueError):
        raise HTTPException(400,"Vui lòng nhập tọa độ hợp lệ.")
    check=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id"}).json()
    if not check: raise HTTPException(404,"Hãy lưu dự án trước khi chụp ảnh.")
    url=f"https://globalsolaratlas.info/detail?c={lat},{lon},10&m=site&s={lat},{lon}"
    labels=["Thông số bức xạ & PV","Bản đồ vị trí công trình","Horizon & Sunpath","Bản đồ tiềm năng PVOUT"]
    try:
        with sync_playwright() as playwright:
            browser=playwright.chromium.launch(headless=True,args=["--no-sandbox","--disable-dev-shm-usage","--disable-gpu","--disable-extensions","--single-process","--no-zygote"])
            page=browser.new_page(viewport={"width":1600,"height":1100},device_scale_factor=1)
            page.goto(url,wait_until="domcontentloaded",timeout=90000)
            try: page.wait_for_load_state("networkidle",timeout=25000)
            except PlaywrightTimeoutError: pass
            page.wait_for_timeout(18000)
            # Chọn các vùng hiển thị lớn có nội dung; loại bỏ header/footer và phần tử trùng lặp.
            candidates=page.evaluate("""() => {
              const nodes=[...document.querySelectorAll('section,article,main,canvas,svg,img,[class*=panel],[class*=card],[class*=widget],[class*=map]')];
              return nodes.map((e,i)=>{const r=e.getBoundingClientRect(),s=getComputedStyle(e);
                return {i,tag:e.tagName,cls:String(e.className||'').slice(0,160),text:(e.innerText||e.getAttribute('aria-label')||'').slice(0,500),
                  x:r.x+scrollX,y:r.y+scrollY,w:r.width,h:r.height,display:s.display,visibility:s.visibility};
              }).filter(x=>x.w>=280&&x.h>=150&&x.w<innerWidth*1.1&&x.h<innerHeight*2.5&&x.display!=='none'&&x.visibility!=='hidden');
            }""")
            # Ưu tiên các vùng có nhãn/đồ họa đặc trưng; fallback dùng 4 vùng lớn khác nhau.
            groups=[
              re.compile(r'PVOUT|GHI|DNI|GTI|irradiation|specific yield|annual',re.I),
              re.compile(r'map|satellite|location|site',re.I),
              re.compile(r'horizon|sunpath|sun path',re.I),
              re.compile(r'PVOUT|potential|photovoltaic',re.I)
            ]
            chosen=[]
            for idx,pat in enumerate(groups):
                ranked=[]
                for c in candidates:
                    if any(abs(c["x"]-q["x"])<30 and abs(c["y"]-q["y"])<30 for q in chosen): continue
                    score=(1200 if pat.search(c["text"]+" "+c["cls"]) else 0)+min(c["w"]*c["h"]/10000,100)
                    if idx==1 and ("map" in c["cls"].lower() or c["tag"] in ("CANVAS","IMG")): score+=80
                    if idx==2 and ("svg"==c["tag"].lower() or "horizon" in c["cls"].lower()): score+=100
                    if score>0: ranked.append((score,c))
                if ranked: chosen.append(max(ranked,key=lambda x:x[0])[1])
            # If Atlas renders as a single dashboard, capture four distinct large visual regions.
            if len(chosen)<4:
                for c in sorted(candidates,key=lambda x:x["w"]*x["h"],reverse=True):
                    if not any(abs(c["x"]-q["x"])<35 and abs(c["y"]-q["y"])<35 for q in chosen):
                        chosen.append(c)
                    if len(chosen)>=4: break
            if len(chosen)<4:
                browser.close()
                raise HTTPException(502,"Global Solar Atlas chưa hiển thị đủ 4 vùng có thể chụp. Có thể trang đã thay đổi giao diện hoặc chặn truy cập tự động.")
            # Screenshot từng vùng qua bounding box, kèm giới hạn kích thước để tránh ảnh quá lớn.
            images=[]
            for i,c in enumerate(chosen[:4]):
                x=max(0,int(c["x"])); y=max(0,int(c["y"])); w=min(int(c["w"]),1600); h=min(int(c["h"]),1100)
                page.evaluate("(y)=>window.scrollTo(0,y)",y)
                page.wait_for_timeout(700)
                box=page.evaluate("(idx)=>{const a=[...document.querySelectorAll('section,article,main,canvas,svg,img,[class*=panel],[class*=card],[class*=widget],[class*=map]')];const e=a[idx];if(!e)return null;const r=e.getBoundingClientRect();return {x:r.x,y:r.y,width:r.width,height:r.height}}",c["i"])
                if not box or box["width"]<1 or box["height"]<1:
                    continue
                data=page.screenshot(type="png",clip={"x":max(0,int(box["x"])),"y":max(0,int(box["y"])),"width":max(1,min(int(box["width"]),1600)),"height":max(1,min(int(box["height"]),1100))})
                images.append(data)
            browser.close()
        if len(images)!=4:
            raise HTTPException(502,f"Chỉ chụp được {len(images)}/4 khu vực trên Global Solar Atlas; chưa lưu bộ ảnh chưa đầy đủ.")
        saved=[]
        for i,data in enumerate(images):
            fid=str(uuid.uuid4()); filename=f"global-solar-atlas-{i+1}.png"; path=storage_path(pid,fid,filename)
            sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(path,safe="/"),data=data,headers={"Content-Type":"image/png","x-upsert":"true"})
            try:
                sb_request("POST",db("uploads"),json={"id":fid,"pid":pid,"name":filename,"path":path},headers={"Prefer":"return=minimal"})
            except Exception:
                sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[path]})
                raise
            saved.append({"id":fid,"name":filename,"label":labels[i]})
        return {"images":saved,"source_url":url,"count":len(saved)}
    except HTTPException: raise
    except Exception as exc:
        print(f"Global Solar Atlas capture error: {type(exc).__name__}: {exc}",flush=True)
        raise HTTPException(502,"Không thể tự động chụp Global Solar Atlas. Kiểm tra Render Logs; có thể Render Free thiếu bộ nhớ hoặc thiếu thư viện trình duyệt.")
    
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
