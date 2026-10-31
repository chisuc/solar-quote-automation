import os, json, uuid, secrets, hmac, mimetypes, asyncio, io, zipfile, re, xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timezone
from urllib.parse import quote
import requests
from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.middleware.sessions import SessionMiddleware
from openpyxl import load_workbook

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

LOGIN="""<!doctype html><html lang='vi'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Đăng nhập | Solar Quote Automation</title><link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%2315352b'/%3E%3Ccircle cx='32' cy='20' r='9' fill='%23f8c548'/%3E%3Cg stroke='%23f8c548' stroke-width='3' stroke-linecap='round'%3E%3Cpath d='M32 5v4M32 31v4M17 20h4M43 20h4M21 9l3 3M40 28l3 3M43 9l-3 3M24 28l-3 3'/%3E%3C/g%3E%3Cpath d='M12 37h40l6 19H6z' fill='%23e8f4ef' stroke='%23168a61' stroke-width='3' stroke-linejoin='round'/%3E%3Cpath d='M25 37l-3 19M39 37l3 19M9 47h46' stroke='%23168a61' stroke-width='2'/%3E%3C/svg%3E"/><style>
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
PAGE="""<!doctype html><html lang='vi'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Solar Quote Automation</title><link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%2315352b'/%3E%3Ccircle cx='32' cy='20' r='9' fill='%23f8c548'/%3E%3Cg stroke='%23f8c548' stroke-width='3' stroke-linecap='round'%3E%3Cpath d='M32 5v4M32 31v4M17 20h4M43 20h4M21 9l3 3M40 28l3 3M43 9l-3 3M24 28l-3 3'/%3E%3C/g%3E%3Cpath d='M12 37h40l6 19H6z' fill='%23e8f4ef' stroke='%23168a61' stroke-width='3' stroke-linejoin='round'/%3E%3Cpath d='M25 37l-3 19M39 37l3 19M9 47h46' stroke='%23168a61' stroke-width='2'/%3E%3C/svg%3E"/><style>*{box-sizing:border-box}body{margin:0;background:#f3f6f5;font:14px Arial;color:#20372e}header{background:#15352b;color:white;padding:18px 24px}main{width:min(96vw,1580px);max-width:none;margin:24px auto;padding:0 14px}
#dashboardView{min-height:70vh}.dash-head{margin:8px 0 18px}.dash-head h1{margin:0}.folder-create{margin-bottom:20px}.folder-create-row{align-items:end}.folder-create-row .field{flex:1;margin:0}.folder-create-row .btn{height:42px;white-space:nowrap}.folder-date-group{margin:24px 0}.folder-date-title{font-size:15px;color:#315b49;font-weight:700;border-bottom:1px solid #dce8e1;padding:0 0 9px;margin-bottom:12px}.folder-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:14px}.folder-card{background:#fff;border:1px solid #dce8e1;border-radius:13px;padding:17px;box-shadow:0 3px 12px #17382c08;transition:.15s}.folder-card:hover{border-color:#9fc8b4;box-shadow:0 8px 22px #17382c12;transform:translateY(-2px)}.folder-icon{font-size:31px;margin-bottom:10px}.folder-card h3{font-size:15px;margin:0 0 6px;overflow-wrap:anywhere}.folder-card .muted{font-size:12px}.folder-actions{display:flex;gap:8px;margin-top:14px}.folder-actions button{flex:1}.editor-toolbar{background:#fff;border:1px solid #dce8e1;border-radius:11px;padding:12px 15px;margin-bottom:18px;align-items:center}.editor-toolbar .muted{font-size:12px}
@media(max-width:650px){.folder-create-row{display:flex;flex-direction:column;align-items:stretch}.folder-create-row .field,.folder-create-row .btn{width:100%}}
.panel{background:white;border:1px solid #dce5e1;border-radius:12px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.field label{display:block;font-size:12px;font-weight:bold;margin-bottom:6px}input,select{width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px}.btn{padding:10px 14px;background:#16845b;color:white;border:0;border-radius:7px;font-weight:bold;cursor:pointer}.alt{background:#e8f1ed;color:#245441}.row{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.project{padding:12px;border:1px solid #dce5e1;border-radius:8px;margin:8px 0;display:flex;justify-content:space-between}.muted{color:#687a73;font-size:12px}.gsa-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}.gsa-card{border:1px solid #dce5e1;border-radius:9px;padding:12px;min-width:0}.gsa-card h4{margin:0 0 12px;color:#245441}.gsa-metrics{width:100%;border-collapse:collapse}.gsa-metrics td{padding:7px 5px;border-bottom:1px solid #e8efeb}.gsa-metrics td:last-child{text-align:right;font-weight:bold}.gsa-sun svg{width:100%;height:auto}.gsa-card a{display:inline-block;margin-top:8px;color:#0879b9}@media(max-width:850px){.compass-field{grid-column:1/-1}.pv-chart-box{grid-column:1/-1;max-width:100%;min-height:250px}.compass-wrap{flex-wrap:wrap;justify-content:flex-start;gap:18px}.compass{width:210px!important;height:210px!important;min-width:210px;min-height:210px;flex-basis:210px}}@media(max-width:700px){.gsa-grid{grid-template-columns:1fr}}.compass-field{grid-column:1/3;min-width:0}.compass{width:220px!important;height:220px!important;min-width:220px;min-height:220px;max-width:none!important;max-height:none!important;aspect-ratio:1/1;flex:0 0 220px;border-radius:50%;box-sizing:border-box}.pv-chart-box{grid-column:3;grid-row:auto;align-self:center;display:flex;flex-direction:column;justify-content:center;min-width:0;width:100%;min-height:300px;margin:0!important;padding:14px 12px;overflow:hidden;border:1px solid #dce5e1;border-radius:10px;background:#fff}.pv-chart-box h4{font-size:13px;line-height:1.3;margin:0 0 8px;color:#245441}.pv-chart-box #pvChartCanvas{width:100%;overflow:hidden;flex:1;display:flex;align-items:center;justify-content:center}.pv-chart-box svg{display:block;width:100%;height:auto;min-width:0!important}.compass-wrap{display:flex;align-items:center;justify-content:center;gap:26px;flex-wrap:nowrap;padding:8px 4px}.compass{width:220px;height:220px;flex:0 0 220px}.compass-info{max-width:285px}.needle{height:68px}.needle i{border-bottom-width:60px}.needle b{top:60px}.compass-info{min-width:150px;flex:1}.compass-info .row{gap:6px}.compass-info .btn{font-size:12px;padding:8px 10px}.compass{width:250px;height:250px;max-width:75vw;max-height:75vw;border:3px solid #9aafa5;border-radius:50%;position:relative;background:radial-gradient(circle,#fff 0 57%,#edf4f0 58% 100%);box-shadow:inset 0 0 0 7px #f7faf8,0 2px 8px #00000012;touch-action:none;user-select:none;cursor:crosshair}.compass:before,.compass:after{content:'';position:absolute;background:#d7e2dc;left:50%;top:10%;width:1px;height:80%;transform:translateX(-50%)}.compass:after{transform:translateX(-50%) rotate(90deg)}.dir{position:absolute;font-size:11px;font-weight:bold;color:#52665d;z-index:2}.north{top:7px;left:50%;transform:translateX(-50%);color:#c0392b}.south{bottom:7px;left:50%;transform:translateX(-50%)}.east{right:8px;top:50%;transform:translateY(-50%)}.west{left:8px;top:50%;transform:translateY(-50%)}.ne{right:34px;top:32px}.se{right:34px;bottom:32px}.sw{left:34px;bottom:32px}.nw{left:34px;top:32px}.needle{position:absolute;z-index:3;left:50%;top:50%;width:8px;height:88px;transform:translate(-50%,-100%) rotate(180deg);transform-origin:50% 100%;pointer-events:none}.needle i{position:absolute;left:0;top:0;width:0;height:0;border-left:4px solid transparent;border-right:4px solid transparent;border-bottom:78px solid #e04435;filter:drop-shadow(0 1px 1px #777)}.needle b{position:absolute;left:1px;top:78px;width:6px;height:10px;background:#667b71;border-radius:0 0 3px 3px}.compass-center{position:absolute;z-index:4;left:50%;top:50%;width:15px;height:15px;border-radius:50%;background:#fff;border:4px solid #315c49;transform:translate(-50%,-50%);pointer-events:none}.compass-info{min-width:230px;flex:1}.compass-info .row{margin:12px 0}@media(max-width:650px){.grid{grid-template-columns:1fr 1fr}}@media(max-width:450px){.grid{grid-template-columns:1fr}}</style></head><body><header><b>☀ SOLAR QUOTE AUTOMATION</b> <button class='btn alt' onclick='logout()' style='float:right'>Đăng xuất</button></header><main>
<section id="dashboardView"><div class="dash-head"><div><h1>Thư mục dự án</h1><p class="muted">Quản lý hồ sơ theo từng ngày, giống thư mục trên máy tính.</p></div></div>
<div class="panel folder-create"><h3>＋ Tạo thư mục mới</h3><div class="row folder-create-row"><div class="field"><label>Tên thư mục</label><input id="newFolderName" placeholder="Ví dụ: Khách hàng Hồng An"></div><div class="field"><label>Ngày tạo</label><input id="newFolderDate" type="date"></div><button class="btn" type="button" onclick="createFolder()">Tạo thư mục</button></div></div>
<div id="folderGroups"></div><div id="folderStatus" class="muted"></div></section>
<section id="folderView" style="display:none"><div class="row editor-toolbar"><button class="btn alt" type="button" onclick="backToFolders()">← Tất cả thư mục</button><div><b id="folderViewTitle">📁 Thư mục</b><div class="muted" id="folderViewDate"></div></div></div><div class="panel"><div class="row" style="justify-content:space-between;align-items:center"><div><h3 style="margin:0">Dự án trong thư mục</h3><p class="muted">Mở dự án đã lưu hoặc bắt đầu tạo dự án mới.</p></div><button class="btn" type="button" onclick="startProject()">＋ Bắt đầu tạo dự án</button></div><div id="folderProjectList" style="margin-top:14px"></div></div></section>
<section id="editorView" style="display:none"><div class="row editor-toolbar"><button class="btn alt" type="button" onclick="backToFolder()">← Quay lại thư mục</button><div><b id="currentFolderTitle">Thư mục dự án</b><div class="muted" id="currentFolderDate"></div></div></div>
<h1>Trung tâm báo giá điện mặt trời</h1><p class='muted'>Website MVP · Dữ liệu lưu trên máy chủ</p>
<div class='panel'><h3>Khách hàng & công trình</h3><div class='grid'><div class='field'><label>Tên dự án</label><input id='projectName'></div><div class='field'><label>Khách hàng / công ty</label><input id='customer'></div><div class='field'><label>Người liên hệ</label><input id='contact'></div><div class='field'><label>Số điện thoại</label><input id='phone'></div><div class='field'><label>Địa điểm công trình</label><input id='address'></div><div class='field'><label>Ngày báo giá</label><input id='date' type='date'></div><div class='field'><label>Ngày hiệu lực báo giá (để trống = +7 ngày)</label><input id='quoteValidUntil' type='date'></div><div class='field'><label>Ngành nghề</label><input id='industry' placeholder='Ví dụ: Sản xuất bao bì'></div><div class='field'><label>Diện tích mái (m²)</label><input id='roof' type='number'></div><div class='field'><label>Bức xạ (kWh/m²/năm)</label><input id='solarRadiation' type='number' step='.1'></div><div class='field'><label>Kết cấu mái</label><input id='roofType' placeholder='Ví dụ: Kèo thép, xà gồ C'></div><div class='field'><label>Điều kiện mái</label><input id='roofCondition' placeholder='Ví dụ: Mái mới, không dột'></div><div class='field'><label>Thói quen dùng điện</label><textarea id='usagePattern' rows='2' style='width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px' placeholder='Ví dụ: Hoạt động 2 ca, từ 7h đến 22h; nghỉ Chủ nhật...'></textarea></div></div></div>
<div class='panel'><h3>Cấu hình hệ thống</h3><div class='grid'><div class='field'><label>Thương hiệu tấm pin</label><input id='panelBrand'></div><div class='field'><label>Công suất tấm pin (Wp)</label><input id='panelW' type='number' value='625'></div><div class='field'><label>Số lượng tấm</label><input id='panelQty' type='number' value='30'></div><div class='field'><label>Công suất PV (kWp)</label><input id='pv' readonly></div><div class='field'><label>Inverter / Model</label><input id='invModel'></div><div class='field'><label>Công suất inverter (kW)</label><input id='invKw' type='number'></div><div class='field'><label>Số pha</label><select id='phase'><option>1 pha</option><option>3 pha</option></select></div><div class='field'><label>Pin lưu trữ / Model</label><input id='batModel'></div><div class='field'><label>Dung lượng mỗi bộ (kWh)</label><input id='batKwh' type='number' step='.01'></div><div class='field'><label>Số lượng bộ pin</label><input id='batQty' type='number' value='0'></div><div class='field'><label>Tổng dung lượng pin (kWh)</label><input id='batTotal' readonly></div></div></div>
<div class='panel'><h3>Điện năng & báo giá</h3><div class='grid'><div class='field'><label>Tiêu thụ tháng (kWh)</label><input id='load' type='number'></div><div class='field'><label>Sản lượng ĐMT năm (kWh)</label><input id='annual' type='number'></div><div class='field'><label>Tự tiêu thụ dự kiến (%)</label><input id='self' type='number' value='80'></div><div class='field'><label>Giá điện trung bình có VAT (đ/kWh)</label><input id='price' type='number'></div><div class='field'><label>Tổng trước VAT (VNĐ)</label><input id='pre' type='number'></div><div class='field'><label>VAT (%)</label><input id='vat' type='number' value='8'></div><div class='field'><label>Tiền VAT</label><input id='vatMoney' readonly></div><div class='field'><label>Tổng sau VAT</label><input id='grand' readonly></div></div></div>
<div class='panel' id='financialPanel'><div class='row' style='justify-content:space-between;align-items:center'><div><h3 style='margin:0'>Phân tích tài chính</h3><p class='muted'>Sử dụng mặc định theo file mẫu. Chỉ điều chỉnh khi cần thiết.</p></div><button class='btn alt' id='financialToggle' type='button' onclick='toggleFinancialAdvanced()'>⚙ Điều chỉnh chuyên sâu PTTC</button></div><div id='financialAdvanced' style='display:none;margin-top:16px'><p class='muted'>Ô để trống sẽ giữ nguyên thông số trong file Excel gốc. Chỉ các ô có nhập giá trị mới được thay đổi.</p><div class='grid'><div class='field'><label>Tỷ lệ lạm phát (%/năm)</label><input id='finInflation' type='number' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Mức tăng giá điện (%/năm)</label><input id='finElectricityEscalation' type='number' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Tỷ lệ tiêu thụ điện (%)</label><input id='finSelfConsumption' type='number' step='0.01' placeholder='Giữ nguyên theo thông số dự án'></div><div class='field'><label>Tỷ lệ vay vốn (%)</label><input id='finLoanRatio' type='number' min='0' max='100' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Thời hạn vay (năm)</label><input id='finLoanTerm' type='number' min='0' step='1' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Lãi suất vay (%/năm)</label><input id='finLoanRate' type='number' min='0' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Thời gian khấu hao (năm)</label><input id='finDepreciation' type='number' min='1' step='1' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Chi phí vận hành (triệu đồng/năm)</label><input id='finOperatingCost' type='number' min='0' step='0.01' placeholder='Giữ nguyên mặc định'></div></div><div class='row'><button class='btn' type='button' onclick='applyFinancialOverrides()'>Thay đổi & tạo báo cáo</button><button class='btn alt' type='button' onclick='resetFinancialOverrides()'>Xóa điều chỉnh</button></div><div id='financialStatus' class='muted' aria-live='polite'></div></div><div class='row' style='align-items:center'><input id='financialTemplateFile' type='file' accept='.xlsx' style='max-width:320px'><button class='btn alt' type='button' onclick='uploadFinancialTemplate()'>Tải mẫu Excel lên hệ thống</button><button class='btn' type='button' onclick='generateFinancialReport()'>Tạo báo cáo PTTC</button></div><div id='financialReportStatus' class='muted' aria-live='polite'></div></div>
<div class='panel' id='presentationPanel'><h3>Xuất báo giá PowerPoint</h3><p class='muted'>Dùng đúng form PowerPoint 14 trang của công ty. Tải file mẫu một lần; các lần sau website sẽ tự dùng lại mẫu đã lưu.</p><div class='row' style='align-items:center'><input id='presentationTemplateFile' type='file' accept='.pptx' style='max-width:340px'><button class='btn alt' type='button' onclick='uploadPresentationTemplate()'>Tải form PowerPoint mẫu</button><button class='btn' type='button' onclick='generatePresentationReport()'>Xuất báo giá PowerPoint</button></div><div id='presentationStatus' class='muted' aria-live='polite'></div></div>
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
<div class='panel'><h3>Hình ảnh phục vụ báo giá PowerPoint</h3><p class='muted'>Tải riêng ảnh hiện trạng mái và ảnh bố trí tấm PV. Ảnh được lưu theo đúng dự án đang mở.</p><div class='grid'><div class='field'><label>1. Hình ảnh mái thực tế</label><input id='roofRealImages' type='file' accept='image/jpeg,image/png,image/webp' multiple><div class='row'><button class='btn' type='button' onclick="uploadProjectImages('roof-real')">Tải ảnh mái thực tế</button></div><div id='roofRealImageList' class='muted'>Chưa có ảnh.</div></div><div class='field'><label>2. Hình ảnh mái bố trí PV</label><input id='roofPvImages' type='file' accept='image/jpeg,image/png,image/webp' multiple><div class='row'><button class='btn' type='button' onclick="uploadProjectImages('roof-pv')">Tải ảnh bố trí PV</button></div><div id='roofPvImageList' class='muted'>Chưa có ảnh.</div></div></div></div><div class='panel'><h3>Hồ sơ đầu vào</h3><input id='files' type='file' multiple><div class='row'><button class='btn' onclick='upload()'>Tải file lên máy chủ</button></div><div id='fileList' class='muted'></div></div>
<div class='panel'><h3>Dự án đã lưu</h3><div class='row'><button class='btn' onclick='save()'>Lưu dự án</button><button class='btn alt' onclick='fresh()'>Dự án mới</button><button class='btn alt' onclick='exportJson()'>Xuất JSON</button></div><div id='projects'>Đang tải...</div><div id='status' class='muted'></div></div></section></main>
<script>
let pid=null,pvwattsResult=null,gsaImages=[],financialReportId=null,currentFolderId=null,currentFolderName='',currentFolderDate='';const ids=['lat','lon','pvSystemSize','moduleType','arrayType','tilt','azimuth','losses','dcAcRatio','invEff','gcr','albedoMode','albedo','projectName','customer','contact','phone','address','date','quoteValidUntil','industry','solarRadiation','roofCondition','usagePattern','roof','roofType','panelBrand','panelW','panelQty','invModel','invKw','phase','batModel','batKwh','batQty','load','annual','self','price','pre','vat'];const $=x=>document.getElementById(x),n=x=>Number($(x).value)||0;
function calc(){let kw=n('panelW')*n('panelQty')/1000,v=n('pre')*n('vat')/100;$('pv').value=kw.toFixed(3);$('batTotal').value=(n('batKwh')*n('batQty')).toFixed(2);$('vatMoney').value=v.toFixed(0);$('grand').value=(n('pre')+v).toFixed(0)}
document.addEventListener('input',e=>{if(e.target.tagName==='INPUT')calc()});
let financialOverrides={};
const financialFields={finInflation:'inflation',finElectricityEscalation:'electricityEscalation',finSelfConsumption:'selfConsumption',finLoanRatio:'loanRatio',finLoanTerm:'loanTerm',finLoanRate:'loanRate',finDepreciation:'depreciationYears',finOperatingCost:'operatingCost'};
function readFinancialOverrides(){let out={};for(const [id,key] of Object.entries(financialFields)){let raw=$(id).value.trim();if(raw!==''){let v=Number(raw);if(!Number.isFinite(v))throw Error('Giá trị '+key+' không hợp lệ.');out[key]=key==='selfConsumption'?v/100:v}}return out}
function showFinancialOverrides(){for(const [id,key] of Object.entries(financialFields))$(id).value=financialOverrides[key]===undefined?'':(key==='selfConsumption'?financialOverrides[key]*100:financialOverrides[key])}
function toggleFinancialAdvanced(){let el=$('financialAdvanced'),open=el.style.display==='none';el.style.display=open?'block':'none';$('financialToggle').textContent=open?'Ẩn điều chỉnh chuyên sâu':'⚙ Điều chỉnh chuyên sâu PTTC'}
async function applyFinancialOverrides(){try{financialOverrides=readFinancialOverrides();$('financialStatus').textContent='Đã áp dụng '+Object.keys(financialOverrides).length+' thông số điều chỉnh. Đang lưu theo dự án...';if(pid){let p=payload();p.financialOverrides=financialOverrides;await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});$('financialStatus').textContent='Đã lưu thông số PTTC. Đang tạo báo cáo theo các thay đổi...';await generateFinancialReport()}else{$('financialStatus').textContent='Hãy lưu dự án trước khi áp dụng điều chỉnh và tạo báo cáo PTTC.'}}catch(e){$('financialStatus').textContent=e.message}}
function resetFinancialOverrides(){financialOverrides={};showFinancialOverrides();$('financialStatus').textContent='Đã xóa các điều chỉnh; sẽ sử dụng thông số mặc định của mẫu PTTC.'}
async function uploadFinancialTemplate(){let f=$('financialTemplateFile').files[0];if(!f){$('financialReportStatus').textContent='Chọn file Excel mẫu trước.';return}let fd=new FormData();fd.append('file',f);try{$('financialReportStatus').textContent='Đang tải mẫu PTTC...';let r=await fetch('/api/financial-template',{method:'POST',body:fd});let d={};try{d=await r.json()}catch(_){}if(!r.ok)throw Error(d.detail||'Không tải được mẫu PTTC');$('financialReportStatus').textContent='Đã lưu mẫu PTTC trên hệ thống.';$('financialTemplateFile').value=''}catch(e){$('financialReportStatus').textContent=e.message}}
async function generateFinancialReport(){if(!pid){$('financialReportStatus').textContent='Hãy lưu dự án trước khi tạo báo cáo PTTC.';return}try{$('financialReportStatus').textContent='Đang tạo file Excel PTTC...';let p=payload();let r=await api('/api/projects/'+pid+'/financial-report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({payload:p})});financialReportId=r.id;let savedPayload=payload();savedPayload.financialReportId=r.id;await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(savedPayload)});$('financialReportStatus').innerHTML='Đã tạo báo cáo: <a href="/api/files/'+r.id+'">'+esc(r.name)+'</a>';await listFiles()}catch(e){$('financialReportStatus').textContent='Không tạo được báo cáo: '+e.message}}
async function uploadPresentationTemplate(){let f=$('presentationTemplateFile').files[0];if(!f){$('presentationStatus').textContent='Chọn file PowerPoint mẫu trước.';return}let fd=new FormData();fd.append('file',f);try{$('presentationStatus').textContent='Đang lưu form PowerPoint mẫu...';let r=await fetch('/api/presentation-template',{method:'POST',body:fd});let d={};try{d=await r.json()}catch(_){}if(!r.ok)throw Error(d.detail||'Không tải được form PowerPoint');$('presentationStatus').textContent='Đã lưu form PowerPoint mẫu trên hệ thống.';$('presentationTemplateFile').value=''}catch(e){$('presentationStatus').textContent=e.message}}
async function generatePresentationReport(){if(!pid){$('presentationStatus').textContent='Hãy lưu dự án trước khi xuất PowerPoint.';return}try{$('presentationStatus').textContent='Đang tạo PowerPoint theo form gốc...';let r=await api('/api/projects/'+pid+'/presentation-report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({payload:payload()})});$('presentationStatus').innerHTML='Đã xuất báo giá: <a href="/api/files/'+r.id+'">'+esc(r.name)+'</a>';await listFiles()}catch(e){$('presentationStatus').textContent='Không xuất được PowerPoint: '+e.message}}


function payload(){let o={};ids.forEach(k=>o[k]=$(k).value);o.financialOverrides=financialOverrides;o.financialReportId=financialReportId;o.pvwattsResult=pvwattsResult;o.globalSolarAtlasResult=globalSolarAtlasResult;o.gsaImages=gsaImages;return o}function fill(o){ids.forEach(k=>{if(o[k]!==undefined)$(k).value=o[k]});financialOverrides=o.financialOverrides||{};financialReportId=o.financialReportId||null;showFinancialOverrides();pvwattsResult=o.pvwattsResult||null;globalSolarAtlasResult=o.globalSolarAtlasResult||null;gsaImages=o.gsaImages||[];renderGSAImages();renderPV();calc()}
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
async function save(){try{let p=payload();p.folderId=currentFolderId;let r=pid?await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)}):await api('/api/projects',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});pid=r.id;$('status').textContent='Đã lưu dự án vào thư mục.';list()}catch(e){$('status').textContent=e.message}}
async function list(){if(!currentFolderId)return;let a=await api('/api/folders/'+currentFolderId+'/projects');let html=a.map(x=>`<div class='project'><span><b>${esc(x.name)}</b><br><small>${esc(x.customer||'')}</small></span><span><button class='btn alt' onclick="openP('${x.id}')">Mở</button> <button class='btn alt' onclick="delP('${x.id}')">Xóa</button></span></div>`).join('')||'<div class="muted">Thư mục chưa có dự án. Nhấn Bắt đầu tạo dự án để thêm hồ sơ đầu tiên.</div>';$('projects').innerHTML=html;$('folderProjectList').innerHTML=html}
async function renderFolders(){let a=await api('/api/folders'),groups={};a.forEach(x=>{let d=x.date||'Chưa phân loại ngày';(groups[d]??=[]).push(x)});$('folderGroups').innerHTML=Object.keys(groups).sort((a,b)=>b.localeCompare(a)).map(d=>`<section class="folder-date-group"><div class="folder-date-title">${esc(d==='Chưa phân loại ngày'?d:new Date(d+'T00:00:00').toLocaleDateString('vi-VN',{weekday:'long',day:'2-digit',month:'2-digit',year:'numeric'}))}</div><div class="folder-grid">${groups[d].map(x=>`<div class="folder-card"><div class="folder-icon">📁</div><h3>${esc(x.name)}</h3><div class="muted">Thư mục dự án · ${esc(x.date||'')}</div><div class="folder-actions"><button class="btn" onclick="openFolder('${x.id}')">Mở thư mục</button><button class="btn alt" onclick="delFolder('${x.id}')">Xóa</button></div></div>`).join('')}</div></section>`).join('')||'<div class="panel muted">Chưa có thư mục. Hãy tạo thư mục đầu tiên để bắt đầu.</div>'}
async function createFolder(){let name=$('newFolderName').value.trim(),date=$('newFolderDate').value;if(!name){$('folderStatus').textContent='Vui lòng nhập tên thư mục.';return}try{let r=await api('/api/folders',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,date})});$('newFolderName').value='';await renderFolders();openFolder(r.id)}catch(e){$('folderStatus').textContent=e.message}}
async function openFolder(id){let a=await api('/api/folders'),f=a.find(x=>x.id===id);if(!f)return;currentFolderId=id;currentFolderName=f.name;currentFolderDate=f.date||'';$('folderViewTitle').textContent='📁 '+f.name;$('folderViewDate').textContent=f.date||'';$('dashboardView').style.display='none';$('editorView').style.display='none';$('folderView').style.display='block';await list()}
function startProject(){pid=null;fresh();$('currentFolderTitle').textContent=$('folderViewTitle').textContent;$('currentFolderDate').textContent=$('folderViewDate').textContent;$('folderView').style.display='none';$('editorView').style.display='block';$('status').textContent='Bắt đầu tạo dự án mới.'}
function backToFolder(){if(pid&&!confirm('Bạn đã lưu các thay đổi chưa? Quay lại thư mục?'))return;$('editorView').style.display='none';$('folderView').style.display='block';pid=null;list()}
function backToFolders(){if(!confirm('Quay lại danh sách tất cả thư mục?'))return;$('folderView').style.display='none';$('editorView').style.display='none';$('dashboardView').style.display='block';currentFolderId=null;pid=null;renderFolders()}
async function delFolder(id){if(!confirm('Xóa thư mục và toàn bộ dự án, file bên trong? Thao tác này không thể hoàn tác.'))return;try{await api('/api/folders/'+id,{method:'DELETE'});await renderFolders()}catch(e){$('folderStatus').textContent=e.message}}
function esc(s){return String(s||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
async function openP(id){let x=await api('/api/projects/'+id);pid=id;currentFolderId=x.payload.folderId||currentFolderId;fill(x.payload);$('folderView').style.display='none';$('dashboardView').style.display='none';$('editorView').style.display='block';$('currentFolderTitle').textContent='📁 '+currentFolderName;$('currentFolderDate').textContent=currentFolderDate;let files=await listFiles();gsaImages=files.filter(v=>String(v.name||'').startsWith('global-solar-atlas-')).sort((a,b)=>a.name.localeCompare(b.name)).map(v=>({id:v.id,name:v.name}));renderGSAImages()}
async function delP(id){if(confirm('Xóa dự án?')){await api('/api/projects/'+id,{method:'DELETE'});if(pid===id)fresh();list()}}
function fresh(){pid=null;financialOverrides={};showFinancialOverrides();$('financialStatus').textContent='';ids.forEach(k=>$(k).value=$(k).defaultValue||'');$('date').value=new Date().toISOString().slice(0,10);$('panelW').value=625;$('panelQty').value=30;$('batQty').value=0;$('vat').value=8;pvwattsResult=null;globalSolarAtlasResult=null;gsaImages=[];renderGSAImages();calc()}
async function upload(){if(!pid){alert('Hãy lưu dự án trước');return}let fd=new FormData();for(let f of $('files').files)fd.append('files',f);let r=await fetch('/api/projects/'+pid+'/files',{method:'POST',body:fd});if(!r.ok){alert('Tải file lỗi');return}$('files').value='';listFiles()}
async function uploadProjectImages(category){if(!pid){alert('Hãy lưu dự án trước khi tải ảnh.');return}let input=category==='roof-real'?$ ('roofRealImages'):$('roofPvImages');let files=[...input.files];if(!files.length){alert('Vui lòng chọn ảnh trước.');return}let fd=new FormData();for(let f of files)fd.append('files',f);fd.append('category',category);try{let r=await fetch('/api/projects/'+pid+'/files',{method:'POST',body:fd});let d={};try{d=await r.json()}catch(_){}if(!r.ok)throw Error(d.detail||'Tải ảnh không thành công');input.value='';await listFiles();$('status').textContent='Đã tải '+d.length+' ảnh lên dự án.'}catch(e){alert(e.message)}}
async function renderProjectImageLists(items){let real=items.filter(x=>String(x.name||'').startsWith('roof-real__')),pv=items.filter(x=>String(x.name||'').startsWith('roof-pv__'));const render=(a)=>a.length?a.map(x=>'<div style="display:flex;align-items:center;gap:8px;margin:7px 0"><img src="/api/files/'+x.id+'" style="width:76px;height:54px;object-fit:cover;border-radius:6px;border:1px solid #dce5e1"><a href="/api/files/'+x.id+'" target="_blank">'+esc(x.name.replace(/^roof-(real|pv)__/,''))+'</a></div>').join(''):'Chưa có ảnh.';$('roofRealImageList').innerHTML=render(real);$('roofPvImageList').innerHTML=render(pv)}

async function listFiles(){if(!pid)return;let a=await api('/api/projects/'+pid+'/files');let ordinary=a.filter(x=>!String(x.name||'').startsWith('roof-real__')&&!String(x.name||'').startsWith('roof-pv__'));$('fileList').innerHTML=ordinary.map(x=>`<div><a href='/api/files/${x.id}'>${esc(x.name)}</a></div>`).join('')||'Chưa có file';renderProjectImageLists(a);return a}
function exportJson(){let b=new Blob([JSON.stringify(payload(),null,2)],{type:'application/json'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='du-an.json';a.click()}
async function logout(){await fetch('/logout',{method:'POST'});location='/login-page'}
$('date').value=new Date().toISOString().slice(0,10);$('newFolderDate').value=new Date().toISOString().slice(0,10);calc();renderFolders().catch(e=>$('folderStatus').textContent=e.message);
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


FINANCIAL_TEMPLATE_PATH="system/financial-analysis-template.xlsx"
FINANCIAL_TEMPLATE_NAME="financial-analysis-template.xlsx"

def restore_excel_connections(original_bytes, generated_bytes):
    """Keep the source workbook's Power Query connection/custom XML package parts."""
    source=zipfile.ZipFile(io.BytesIO(original_bytes))
    generated=zipfile.ZipFile(io.BytesIO(generated_bytes))
    out=io.BytesIO()
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as target:
        generated_names=set(generated.namelist())
        for item in generated.infolist():
            if item.filename in ("xl/_rels/workbook.xml.rels","[Content_Types].xml"): continue
            target.writestr(item,generated.read(item.filename))
        preserved=["xl/connections.xml","customXml/item1.xml","customXml/itemProps1.xml","customXml/_rels/item1.xml.rels"]
        for name in preserved:
            if name in source.namelist() and name not in generated_names:
                target.writestr(name,source.read(name))
        rel_name="xl/_rels/workbook.xml.rels"
        rel_root=ET.fromstring(generated.read(rel_name))
        rel_ns="http://schemas.openxmlformats.org/package/2006/relationships"
        ET.register_namespace("",rel_ns)
        existing_targets={x.attrib.get("Target","") for x in rel_root}
        next_id=1
        used={x.attrib.get("Id","") for x in rel_root}
        def new_id():
            nonlocal next_id
            while "rId"+str(next_id) in used: next_id+=1
            value="rId"+str(next_id); used.add(value); next_id+=1
            return value
        relations=[
            ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections","connections.xml"),
            ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml","../customXml/item1.xml")
        ]
        for typ,target in relations:
            if target in existing_targets: continue
            part=("xl/"+target).replace("xl/../","")
            if part not in source.namelist(): continue
            ET.SubElement(rel_root,"{"+rel_ns+"}Relationship",{"Id":new_id(),"Type":typ,"Target":target})
        target.writestr(rel_name,ET.tostring(rel_root,encoding="utf-8",xml_declaration=True))
        ct_name="[Content_Types].xml"
        ct_root=ET.fromstring(generated.read(ct_name))
        ct_ns="http://schemas.openxmlformats.org/package/2006/content-types"
        ET.register_namespace("",ct_ns)
        overrides={x.attrib.get("PartName") for x in ct_root}
        needed={
          "/xl/connections.xml":"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
          "/customXml/itemProps1.xml":"application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
        }
        for part,ctype in needed.items():
            if part not in overrides and part.lstrip("/") in source.namelist():
                ET.SubElement(ct_root,"{"+ct_ns+"}Override",{"PartName":part,"ContentType":ctype})
        target.writestr(ct_name,ET.tostring(ct_root,encoding="utf-8",xml_declaration=True))
    source.close(); generated.close()
    return out.getvalue()

@app.post("/api/financial-template")
async def upload_financial_template(request:Request,file:UploadFile=File(...)):
    auth(request)
    name=Path(file.filename or "").name
    if Path(name).suffix.lower()!=".xlsx":
        raise HTTPException(400,"Mẫu PTTC phải là file .xlsx.")
    data=await file.read()
    if len(data)>10*1024*1024:
        raise HTTPException(413,"File mẫu tối đa 10 MB.")
    try:
        check=load_workbook(io.BytesIO(data),read_only=True,data_only=False)
        required={"THONG SO DAU VAO","present","draft ","16,66"}
        missing=required-set(check.sheetnames)
        check.close()
        if missing: raise ValueError("Thiếu sheet: "+", ".join(sorted(missing)))
    except Exception as exc:
        raise HTTPException(400,"File Excel mẫu không hợp lệ hoặc thiếu sheet cần thiết: "+str(exc)[:250])
    sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(FINANCIAL_TEMPLATE_PATH,safe="/"),data=data,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","x-upsert":"true"})
    return {"ok":True,"name":FINANCIAL_TEMPLATE_NAME}

@app.post("/api/projects/{pid}/financial-report")
async def create_financial_report(pid:str,request:Request):
    auth(request)
    body=await request.json()
    project=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id,payload"}).json()
    if not project or (project[0].get("payload") or {}).get("_recordType")=="folder":
        raise HTTPException(404,"Không tìm thấy dự án.")
    saved=project[0].get("payload") or {}
    p=body.get("payload") or saved
    overrides=(p.get("financialOverrides") if "financialOverrides" in p else saved.get("financialOverrides")) or {}
    try:
        original=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(FINANCIAL_TEMPLATE_PATH,safe="/")).content
    except HTTPException as exc:
        if exc.status_code==404: raise HTTPException(409,"Chưa có file mẫu PTTC. Hãy tải file Excel mẫu lên trước.")
        raise
    try:
        wb=load_workbook(io.BytesIO(original),data_only=False,keep_links=True)
        inp=wb["THONG SO DAU VAO"]; present=wb["present"]; draft=wb["draft "]
        def number(key, fallback=None):
            value=p.get(key)
            if value is None or value=="": return fallback
            try: return float(value)
            except (TypeError,ValueError): raise HTTPException(400,f"Thông số {key} không hợp lệ.")
        pv=number("pv")
        if pv is None:
            pw=number("panelW"); qty=number("panelQty")
            if pw is not None and qty is not None: pv=pw*qty/1000
        if pv is not None and pv>0: inp["C3"]=pv
        pre=number("pre")
        if pre is not None: inp["C4"]=pre
        price=number("price")
        if price is not None: inp["C7"]=price
        self_ratio=overrides.get("selfConsumption")
        if self_ratio is None:
            raw_self=number("self")
            self_ratio=raw_self/100 if raw_self is not None else None
        if self_ratio is not None:
            if not 0<=float(self_ratio)<=1: raise HTTPException(400,"Tỷ lệ tiêu thụ điện phải từ 0 đến 100%.")
            inp["C8"]=float(self_ratio)
        monthly=(p.get("pvwattsResult") or {}).get("monthly") or []
        if len(monthly)==12:
            for i,value in enumerate(monthly,11):
                inp[f"C{i}"]=float(value)
        if p.get("projectName"): present["C7"]=str(p["projectName"])
        if p.get("address"): present["D8"]=str(p["address"])
        roof=number("roof")
        if roof is not None: present["D11"]=roof
        override_map={
          "inflation":("present","S46",0.0,1.0),
          "electricityEscalation":("present","S47",0.0,1.0),
          "loanRatio":("present","M49",0.0,100.0),
          "loanTerm":("present","P49",0.0,30.0),
          "loanRate":("present","S49",0.0,100.0),
          "depreciationYears":("draft ","B14",1.0,50.0),
          "operatingCost":("draft ","B15",0.0,1000000000.0)
        }
        for key,(sheet,cell,low,high) in override_map.items():
            if key not in overrides or overrides[key] in (None,""): continue
            value=float(overrides[key])
            if not low<=value<=high: raise HTTPException(400,f"Thông số {key} nằm ngoài giới hạn cho phép.")
            if key in ("inflation","electricityEscalation","loanRate"): value/=100
            if key=="loanRatio": value/=100
            wb[sheet][cell]=value
        # Force Excel to recalculate the original formulas when the generated workbook is opened.
        wb.calculation.fullCalcOnLoad=True
        wb.calculation.forceFullCalc=True
        wb.calculation.calcMode="auto"
        out=io.BytesIO(); wb.save(out); wb.close()
        generated=_recalculate_xlsx(out.getvalue())
        generated=restore_excel_connections(original,generated)
    except HTTPException: raise
    except Exception as exc:
        print(f"Financial report generation error: {type(exc).__name__}: {exc}",flush=True)
        raise HTTPException(500,"Không thể tạo báo cáo PTTC: "+str(exc)[:300])
    project_name=str(p.get("projectName") or p.get("customer") or "Du an").strip()
    safe_name=re.sub(r"[^A-Za-z0-9._-]+","_",project_name).strip("._")[:60] or "Du_an"
    filename=f"PTTC_{safe_name}.xlsx"
    fid=str(uuid.uuid4()); path=storage_path(pid,fid,filename)
    sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(path,safe="/"),data=generated,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","x-upsert":"true"})
    try:
        sb_request("POST",db("uploads"),json={"id":fid,"pid":pid,"name":filename,"path":path},headers={"Prefer":"return=minimal"})
    except Exception:
        sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[path]})
        raise
    return {"id":fid,"name":filename,"recalculation":"Excel sẽ tính lại công thức khi mở file."}


PRESENTATION_TEMPLATE_PATH="system/solar-quote-template.pptx"
PRESENTATION_TEMPLATE_NAME="solar-quote-template.pptx"

@app.post("/api/presentation-template")
async def upload_presentation_template(request:Request,file:UploadFile=File(...)):
    auth(request)
    name=Path(file.filename or "").name
    if Path(name).suffix.lower()!=".pptx":
        raise HTTPException(400,"Form PowerPoint phải là file .pptx.")
    data=await file.read()
    if len(data)>30*1024*1024:
        raise HTTPException(413,"File PowerPoint mẫu tối đa 30 MB.")
    try:
        from pptx import Presentation
        check=Presentation(io.BytesIO(data))
        if len(check.slides)<14:
            raise ValueError("Form cần có ít nhất 14 slide theo mẫu đã thống nhất.")
    except Exception as exc:
        raise HTTPException(400,"File PowerPoint không hợp lệ hoặc không đủ 14 slide: "+str(exc)[:250])
    sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(PRESENTATION_TEMPLATE_PATH,safe="/"),data=data,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.presentationml.presentation","x-upsert":"true"})
    return {"ok":True,"name":PRESENTATION_TEMPLATE_NAME,"slides":14}

def _recalculate_xlsx(data):
    """Recalculate workbook formulas with LibreOffice so data_only reads contain current values."""
    import shutil, subprocess, tempfile
    office=shutil.which("libreoffice") or shutil.which("soffice")
    if not office:
        raise HTTPException(503,"Máy chủ chưa có LibreOffice để tính lại công thức PTTC.")
    with tempfile.TemporaryDirectory() as td:
        src=Path(td)/"input.xlsx"
        outdir=Path(td)/"recalc"
        outdir.mkdir()
        src.write_bytes(data)
        proc=subprocess.run([office,"-env:UserInstallation=file://"+str(Path(td)/"lo-profile"),"--headless","--convert-to","xlsx","--outdir",str(outdir),str(src)],capture_output=True,text=True,timeout=120)
        result=outdir/"input.xlsx"
        if proc.returncode!=0 or not result.exists():
            raise HTTPException(500,"Không thể tính lại công thức PTTC: "+(proc.stderr or proc.stdout or "LibreOffice error")[:250])
        return result.read_bytes()


def _ppt_set_paragraph_text(paragraph,text):
    """Replace paragraph text while retaining its paragraph and character styling."""
    from copy import deepcopy
    text=str(text)
    runs=list(paragraph.runs)
    if not runs:
        paragraph.add_run().text=text
        return
    # Group adjacent runs with identical character formatting, so mixed bold/regular
    # emphasis and font sizes in the source template are retained.
    groups=[]
    for run in runs:
        f=run.font
        sig=(f.name, f.size, f.bold, f.italic, f.underline,
             f.color.type if f.color else None,
             f.color.rgb if f.color and f.color.type and f.color.type==1 else None)
        if groups and groups[-1][0]==sig:
            groups[-1][1].append(run)
        else:
            groups.append((sig,[run]))
    weights=[max(1,sum(len(r.text) for r in group)) for _,group in groups]
    total=sum(weights)
    pos=0
    for i,((_,group),weight) in enumerate(zip(groups,weights)):
        end=len(text) if i==len(groups)-1 else pos+round(len(text)*weight/total)
        chunk=text[pos:end] if i==len(groups)-1 else text[pos:end]
        pos=end
        group[0].text=chunk
        for extra in group[1:]:
            extra.text=""
    # If replacement is empty, retain the paragraph but clear its visible runs.
    if not text:
        for run in runs: run.text=""

def _ppt_set_shape_text(shape,text):
    """Set newline-separated content without flattening the template's formatting."""
    from copy import deepcopy
    tf=shape.text_frame
    lines=str(text).split("\n")
    paragraphs=list(tf.paragraphs)
    if not paragraphs:
        tf.text=str(text)
        return
    for i,para in enumerate(paragraphs[:len(lines)]):
        _ppt_set_paragraph_text(para, lines[i])
    # Remove unused source paragraphs so fewer replacement lines do not leave
    # large blank gaps in the generated slide.
    for para in paragraphs[len(lines):]:
        para._element.getparent().remove(para._element)
    # If the replacement needs more lines than the template, clone the last
    # paragraph's formatting before adding the additional content.
    for line in lines[len(paragraphs):]:
        para=tf.add_paragraph()
        if paragraphs:
            source=paragraphs[-1]
            para.level=source.level
            para.alignment=source.alignment
            if source.runs:
                para.add_run()._r.get_or_add_rPr()
                if source.runs[0]._r.rPr is not None:
                    para.runs[0]._r.rPr=deepcopy(source.runs[0]._r.rPr)
        _ppt_set_paragraph_text(para,line)

def _ppt_set_cell_text(cell,text):
    """Update a table cell while keeping its original paragraph/run styles."""
    _ppt_set_shape_text(cell,text)

def _ppt_monthly_production_chart(workbook_bytes):
    """Render monthly production data used by the PTTC sheet's production chart."""
    from PIL import Image, ImageDraw, ImageFont
    wb=load_workbook(io.BytesIO(workbook_bytes),data_only=True,read_only=True)
    try:
        ws=wb["draft "]
        vals=[float(ws[f"G{i}"].value or 0) for i in range(53,65)]
    finally:
        wb.close()
    import math
    W,H=1300,1000
    im=Image.new("RGB",(W,H),"white")
    d=ImageDraw.Draw(im)
    try:
        font=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",22)
        small=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",18)
        bold=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",22)
    except Exception:
        font=small=bold=ImageFont.load_default()
    left,right,top,bottom=115,45,55,100
    ymax=max(500,math.ceil(max(vals)/500)*500) if vals else 3000
    for k in range(7):
        v=ymax*k/6
        y=H-bottom-int((H-top-bottom)*k/6)
        d.line((left,y,W-right,y),fill=(218,218,218),width=2)
        d.text((left-15,y-13),f"{v:,.0f}",font=small,fill=(40,40,40),anchor="ra")
    step=(W-left-right)/12
    bw=step*.58
    for i,v in enumerate(vals):
        x=left+i*step+(step-bw)/2
        bh=(H-top-bottom)*v/ymax
        y=H-bottom-bh
        d.rectangle((x,y,x+bw,H-bottom),fill=(247,151,68))
        d.text((x+bw/2,y-30),f"{v:.2f}",font=small,fill=(35,35,35),anchor="mm")
        d.text((x+bw/2,H-bottom+28),str(i+1),font=small,fill=(35,35,35),anchor="mm")
    d.text((left,H-42),"Tháng",font=font,fill=(20,20,20))
    d.text((22,top),"kWh",font=bold,fill=(20,20,20))
    out=io.BytesIO(); im.save(out,format="PNG",optimize=True); return out.getvalue()

def _ppt_replace_picture(slide,shape,image_bytes):
    image_part,rid=slide.part.get_or_add_image_part(io.BytesIO(image_bytes))
    blip=shape._element.blipFill.blip
    from pptx.oxml.ns import qn
    blip.set(qn("r:embed"),rid)

def _ppt_photo_collage(items, target_ratio):
    from PIL import Image, ImageOps, ImageDraw
    thumbs=[]
    for raw in items[:4]:
        try:
            im=Image.open(io.BytesIO(raw)).convert("RGB")
            thumbs.append(im)
        except Exception: pass
    if not thumbs: return None
    if len(thumbs)==1:
        out=ImageOps.fit(thumbs[0],(1600,max(400,int(1600/max(target_ratio,.2)))),method=Image.Resampling.LANCZOS)
    else:
        cols=2; rows=(len(thumbs)+1)//2
        W=1600; H=max(500,int(W/max(target_ratio,.2)))
        gap=12; cw=(W-gap*(cols+1))//cols; ch=(H-gap*(rows+1))//rows
        out=Image.new("RGB",(W,H),"white")
        for i,im in enumerate(thumbs):
            x=gap+(i%cols)*(cw+gap); y=gap+(i//cols)*(ch+gap)
            tile=ImageOps.fit(im,(cw,ch),method=Image.Resampling.LANCZOS)
            out.paste(tile,(x,y))
    b=io.BytesIO(); out.save(b,format="JPEG",quality=92); return b.getvalue()

@app.post("/api/projects/{pid}/presentation-report")
async def create_presentation_report(pid:str,request:Request):
    auth(request)
    body=await request.json()
    rows=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id,payload"}).json()
    if not rows or (rows[0].get("payload") or {}).get("_recordType")=="folder":
        raise HTTPException(404,"Không tìm thấy dự án.")
    saved=rows[0].get("payload") or {}
    p=body.get("payload") or saved
    try:
        original=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(PRESENTATION_TEMPLATE_PATH,safe="/")).content
    except HTTPException as exc:
        detail=str(exc.detail or "").lower()
        if exc.status_code in (400,404) or "object not found" in detail or "not found" in detail:
            raise HTTPException(409,"Chưa tìm thấy form PowerPoint mẫu trên máy chủ. Hãy chọn file .pptx 14 trang, nhấn 'Tải form PowerPoint mẫu', đợi thông báo tải thành công rồi mới nhấn 'Xuất báo giá PowerPoint'.")
        raise
    try:
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        prs=Presentation(io.BytesIO(original))
        if len(prs.slides)<14: raise ValueError("Form PowerPoint phải có ít nhất 14 slide.")
        def val(k,default=""):
            x=p.get(k)
            return default if x is None else str(x)
        def num(k,default=0):
            try: return float(p.get(k) or default)
            except (TypeError,ValueError): return float(default)
        def fmt(x,d=2):
            return f"{x:,.{d}f}".replace(",","X").replace(".",",").replace("X",".")
        project=val("projectName",val("customer","DỰ ÁN ĐIỆN MẶT TRỜI"))
        customer=val("customer",project)
        pv=num("pv",num("panelW")*num("panelQty")/1000)
        battery=num("batKwh")*num("batQty")
        monthly_load=num("load")
        annual=num("annual",(p.get("pvwattsResult") or {}).get("annual",0))
        self_use=num("self")
        price=num("price")
        total=num("pre")*(1+num("vat",8)/100)
        date=val("date",datetime.now().astimezone().strftime("%Y-%m-%d"))
        roof=num("roof")
        def text_shape(slide,index,text):
            if index<len(slide.shapes):
                sh=slide.shapes[index]
                if sh.has_text_frame: _ppt_set_shape_text(sh,text)
        # Slide 1: replace investor name and quote dates only.
        cover=prs.slides[0]
        if len(cover.shapes)>7 and cover.shapes[7].has_text_frame:
            paras=cover.shapes[7].text_frame.paragraphs
            if len(paras)>4 and paras[4].runs:
                paras[4].runs[0].text=customer
                for run in paras[4].runs[1:]: run.text=""
        quote_date=date or datetime.now().astimezone().strftime("%Y-%m-%d")
        try:
            d0=datetime.strptime(quote_date,"%Y-%m-%d")
            quote_date_display=d0.strftime("%d/%m/%Y")
            valid_raw=val("quoteValidUntil").strip()
            valid_date=datetime.strptime(valid_raw,"%Y-%m-%d") if valid_raw else d0+__import__("datetime").timedelta(days=7)
            valid_date_display=valid_date.strftime("%d/%m/%Y")
        except ValueError:
            raise HTTPException(400,"Ngày báo giá hoặc ngày hiệu lực không hợp lệ.")
        if len(cover.shapes)>8 and cover.shapes[8].has_text_frame:
            para=cover.shapes[8].text_frame.paragraphs[0]
            if para.runs:
                para.runs[0].text=f"TP.HCM ngày {quote_date_display}, hiệu lực đến ngày {valid_date_display}"
                for run in para.runs[1:]: run.text=""
        # Slide 3: update only self-consumption and 30-year savings from PTTC present!R113.
        slide3=prs.slides[2]
        percent_text=fmt(self_use,1).rstrip("0").rstrip(",")+"%"
        if len(slide3.shapes)>7 and slide3.shapes[7].has_text_frame:
            numeric_runs=[run for run in slide3.shapes[7].text_frame.paragraphs[2].runs if re.search(r"\d",run.text)]
            if numeric_runs:
                numeric_runs[0].text=percent_text
                for run in numeric_runs[1:]: run.text=""
        report_id=str(p.get("financialReportId") or "").strip()
        pttc_rows=sb_request("GET",db("uploads"),params={"id":f"eq.{report_id}","pid":f"eq.{pid}","select":"name,path"}).json() if report_id else []
        if not pttc_rows:
            raise HTTPException(409,"Chưa xác định được file PTTC của dự án. Hãy tạo báo cáo PTTC trước khi xuất PowerPoint.")
        pttc_bytes=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(pttc_rows[0]["path"],safe="/")).content
        pttc_wb=load_workbook(io.BytesIO(pttc_bytes),data_only=True,read_only=True)
        try:
            accumulated=pttc_wb["present"]["R113"].value
        finally:
            pttc_wb.close()
        if accumulated is None:
            raise HTTPException(409,"Ô present!R113 chưa có kết quả. Hãy tạo lại báo cáo PTTC rồi xuất PowerPoint.")
        savings_text=f"{float(accumulated):.1f}".replace(".",",")
        pttc_read=load_workbook(io.BytesIO(pttc_bytes),data_only=True,read_only=True)
        try:
            annual_pttc=sum(float(pttc_read["draft "][f"G{i}"].value or 0) for i in range(53,65))
        finally:
            pttc_read.close()
        if len(slide3.shapes)>8 and slide3.shapes[8].has_text_frame:
            paras=slide3.shapes[8].text_frame.paragraphs
            if len(paras)>1 and paras[1].runs:
                paras[1].runs[0].text="LÊN ĐẾN "+savings_text+" TỶ ĐỒNG"
                for run in paras[1].runs[1:]: run.text=""
        # Slide 4: project data table
        s4=prs.slides[3]
        for sh in s4.shapes:
            if sh.has_table:
                table=sh.table
                mapping={
                    "địa điểm":val("address"),"ngành nghề":val("industry"),"quốc gia":"Việt Nam",
                    "điện năng tiêu thụ":(fmt(monthly_load,0)+" kWh / tháng") if monthly_load else "",
                    "sản lượng tiêu thụ ước tính":(fmt(monthly_load*12,0)+" kWh/ năm") if monthly_load else "",
                    "điện năng":(fmt(monthly_load*12,0)+" kWh / năm") if monthly_load else "",
                    "giá điện trung bình":(fmt(price,0)+" VNĐ/kWh") if price else "",
                    "đơn giá điện":(fmt(price,0)+" VNĐ/kWh") if price else "",
                    "diện tích mái":(fmt(roof,0)+" m²") if roof else "",
                    "bức xạ":(fmt(num("solarRadiation"),1)+" kWh/m²/năm") if num("solarRadiation") else "",
                    "kết cấu":val("roofType"),"điều kiện mái":val("roofCondition")
                }
                for row_index in range(1, len(table.rows)):
                    row=table.rows[row_index]
                    label=(row.cells[0].text or "").strip().lower()
                    for key,value in mapping.items():
                        if key in label and value:
                            _ppt_set_cell_text(row.cells[1],value); break
        # Slide 5: sizing, tariff and usage
        text_shape(prs.slides[4],7,f"Đơn giá điện trung bình :\n{fmt(price,0)} VNĐ/kWh\n(bao gồm VAT)")
        text_shape(prs.slides[4],9,"Thói quen dùng điện: "+val("usagePattern","Chưa cập nhật"))
        text_shape(prs.slides[4],11,f"Hệ thống đề xuất:\nCông suất: {fmt(pv)} kWp\nDung lượng lưu trữ: {fmt(battery)} kWh\nTỉ lệ tiêu thụ điện ước tính: {fmt(self_use,1)}%")
        text_shape(prs.slides[4],14,f"Số điện:\n~{fmt(monthly_load,0)} kWh/tháng")
        # Slide 6: radiation/production summary; keep original chart/panels if no new source imagery
        text_shape(prs.slides[5],9,f"Biểu đồ sản lượng điện mặt trời dự kiến sản xuất của hệ thống {fmt(pv)} kWp\nTổng sản lượng 1 năm: {fmt(annual_pttc,0)} kWh")
        # Slides 7 onward are intentionally kept unchanged for this phase.
        # Replace the actual-roof and PV-layout photos on slide 4; keep original geometry/cropping.
        uploads=sb_request("GET",db("uploads"),params={"pid":f"eq.{pid}","select":"id,name,path","order":"name.asc"}).json()
        def image_group(prefix):
            selected=[u for u in uploads if str(u.get("name","")).startswith(prefix+"__")][:4]
            raw=[]
            for item in selected:
                try:
                    obj=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(item["path"],safe="/"))
                except HTTPException as exc:
                    raise HTTPException(409,f"Không tìm thấy file ảnh '{item.get('name','')}' trong kho lưu trữ. Hãy tải lại ảnh ở đúng nhóm rồi thử xuất lại.")
                raw.append(obj.content)
            return raw
        real=image_group("roof-real"); layout=image_group("roof-pv")
        if not real or not layout:
            raise HTTPException(409,"Vui lòng tải ít nhất 1 ảnh mái thực tế và 1 ảnh mái bố trí PV trước khi xuất PowerPoint.")
        def replace_at(slide_index,shape_index,images):
            if not images:return
            slide=prs.slides[slide_index]
            if shape_index>=len(slide.shapes):return
            shape=slide.shapes[shape_index]
            if shape.shape_type!=MSO_SHAPE_TYPE.PICTURE:return
            ratio=shape.width/max(shape.height,1)
            collage=_ppt_photo_collage(images,ratio)
            if collage: _ppt_replace_picture(slide,shape,collage)
        replace_at(3,8,real); replace_at(3,9,layout)
        # Slide 6: production chart rendered from the monthly series used by PTTC.
        if len(prs.slides[5].shapes)>10 and prs.slides[5].shapes[10].shape_type==MSO_SHAPE_TYPE.PICTURE:
            _ppt_replace_picture(prs.slides[5],prs.slides[5].shapes[10],_ppt_monthly_production_chart(pttc_bytes))
        # Four Global Solar Atlas captures occupy the four original picture frames on slide 6.
        for image_no,shape_index in enumerate((11,12,13,14),start=1):
            found=[u for u in uploads if str(u.get("name",""))==f"global-solar-atlas-{image_no}.png"]
            if found and shape_index<len(prs.slides[5].shapes):
                obj=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(found[-1]["path"],safe="/"))
                _ppt_replace_picture(prs.slides[5],prs.slides[5].shapes[shape_index],obj.content)
        # Save the deck while preserving the original theme, master, layout, and all static slides.
        out=io.BytesIO(); prs.save(out); generated=out.getvalue()
    except HTTPException: raise
    except Exception as exc:
        print(f"PowerPoint report generation error: {type(exc).__name__}: {exc}",flush=True)
        raise HTTPException(500,"Không thể tạo báo giá PowerPoint: "+str(exc)[:350])
    safe=re.sub(r"[^A-Za-z0-9._-]+","_",project).strip("._")[:60] or "Du_an"
    filename=f"Bao_gia_{safe}.pptx"; fid=str(uuid.uuid4()); path=storage_path(pid,fid,filename)
    sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(path,safe="/"),data=generated,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.presentationml.presentation","x-upsert":"true"})
    try:
        sb_request("POST",db("uploads"),json={"id":fid,"pid":pid,"name":filename,"path":path},headers={"Prefer":"return=minimal"})
    except Exception:
        sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[path]}); raise
    return {"id":fid,"name":filename,"slides":len(prs.slides),"photos":{"roof_real":len(real),"roof_pv":len(layout)}}

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
    
@app.get("/api/folders")
def folders(request:Request):
    auth(request)
    rows=sb_request("GET",db("projects"),params={"select":"id,name,payload,updated","order":"updated.desc"}).json()
    return [{"id":r["id"],"name":(r.get("payload") or {}).get("folderName") or r.get("name") or "Thư mục","date":(r.get("payload") or {}).get("folderDate",""),"updated":r.get("updated")} for r in rows if (r.get("payload") or {}).get("_recordType")=="folder"]

@app.post("/api/folders")
async def create_folder(request:Request):
    auth(request); p=await request.json()
    name=str(p.get("name") or "").strip()
    folder_date=str(p.get("date") or datetime.now().astimezone().date().isoformat())
    if not name: raise HTTPException(400,"Vui lòng nhập tên thư mục.")
    if len(name)>120: raise HTTPException(400,"Tên thư mục tối đa 120 ký tự.")
    fid=str(uuid.uuid4()); now=datetime.now(timezone.utc).isoformat()
    payload={"_recordType":"folder","folderName":name,"folderDate":folder_date}
    sb_request("POST",db("projects"),json={"id":fid,"name":name,"customer":"","payload":payload,"updated":now},headers={"Prefer":"return=minimal"})
    return {"id":fid,"name":name,"date":folder_date}

@app.get("/api/folders/{fid}/projects")
def folder_projects(fid:str,request:Request):
    auth(request)
    rows=sb_request("GET",db("projects"),params={"select":"id,name,customer,updated,payload","order":"updated.desc"}).json()
    return [{"id":r["id"],"name":r.get("name"),"customer":r.get("customer"),"updated":r.get("updated")} for r in rows if (r.get("payload") or {}).get("folderId")==fid and (r.get("payload") or {}).get("_recordType")!="folder"]

@app.delete("/api/folders/{fid}")
def delete_folder(fid:str,request:Request):
    auth(request)
    rows=sb_request("GET",db("projects"),params={"select":"id,payload"}).json()
    child=[r["id"] for r in rows if (r.get("payload") or {}).get("folderId")==fid and (r.get("payload") or {}).get("_recordType")!="folder"]
    for pid in child:
        ups=sb_request("GET",db("uploads"),params={"pid":f"eq.{pid}","select":"path"}).json()
        if ups: sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[x["path"] for x in ups]})
        sb_request("DELETE",db("uploads"),params={"pid":f"eq.{pid}"})
        sb_request("DELETE",db("projects"),params={"id":f"eq.{pid}"})
    sb_request("DELETE",db("projects"),params={"id":f"eq.{fid}"})
    return {"ok":True,"deleted_projects":len(child)}

@app.get("/api/projects")
def projects(request:Request):
    auth(request)
    r=sb_request("GET",db("projects"),params={"select":"id,name,customer,updated,payload","order":"updated.desc"})
    return [x for x in r.json() if (x.get("payload") or {}).get("_recordType")!="folder"]

@app.post("/api/projects")
async def create(request:Request):
    auth(request); p=await request.json()
    folder_id=p.get("folderId")
    if not folder_id: raise HTTPException(400,"Vui lòng mở thư mục trước khi tạo dự án.")
    frows=sb_request("GET",db("projects"),params={"id":f"eq.{folder_id}","select":"id,payload"}).json()
    if not frows or (frows[0].get("payload") or {}).get("_recordType")!="folder": raise HTTPException(404,"Không tìm thấy thư mục.")
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
    if not rows or (rows[0].get("payload") or {}).get("_recordType")=="folder": raise HTTPException(404,"Không tìm thấy dự án")
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
async def upload(pid:str,request:Request,files:list[UploadFile]=File(...),category:str=Form("")):
    auth(request)
    check=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id"}).json()
    if not check: raise HTTPException(404,"Không tìm thấy dự án")
    out=[]; allowed={".pdf",".xlsx",".xls",".csv",".jpg",".jpeg",".png",".webp",".doc",".docx"}
    if category and category not in ("roof-real","roof-pv"): raise HTTPException(400,"Nhóm ảnh không hợp lệ.")
    for f in files:
        name=Path(f.filename or "file").name; ext=Path(name).suffix.lower()
        if category:
            if ext not in (".jpg",".jpeg",".png",".webp"): raise HTTPException(400,"Ảnh mái chỉ hỗ trợ JPG, PNG hoặc WEBP.")
            name=category+"__"+name
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
