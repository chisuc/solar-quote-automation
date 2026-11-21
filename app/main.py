import os, json, uuid, secrets, hmac, mimetypes, asyncio, io, zipfile, re, time, xml.etree.ElementTree as ET
from pathlib import Path
from typing import List
from datetime import datetime, timezone
from urllib.parse import quote
import requests
from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from starlette.middleware.sessions import SessionMiddleware
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageOps
import pytesseract
from pdf2image import convert_from_bytes

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
.panel{background:white;border:1px solid #dce5e1;border-radius:12px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.field label{display:block;font-size:12px;font-weight:bold;margin-bottom:6px}input,select{width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px}.btn{padding:10px 14px;background:#16845b;color:white;border:0;border-radius:7px;font-weight:bold;cursor:pointer}.alt{background:#e8f1ed;color:#245441}.row{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.project{padding:12px;border:1px solid #dce5e1;border-radius:8px;margin:8px 0;display:flex;justify-content:space-between}.muted{color:#687a73;font-size:12px}.gsa-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}.gsa-card{border:1px solid #dce5e1;border-radius:9px;padding:12px;min-width:0}.gsa-card h4{margin:0 0 12px;color:#245441}.gsa-metrics{width:100%;border-collapse:collapse}.gsa-metrics td{padding:7px 5px;border-bottom:1px solid #e8efeb}.gsa-metrics td:last-child{text-align:right;font-weight:bold}.gsa-sun svg{width:100%;height:auto}.gsa-card a{display:inline-block;margin-top:8px;color:#0879b9}@media(max-width:850px){.compass-field{grid-column:1/-1}.pv-chart-box{grid-column:1/-1;max-width:100%;min-height:250px}.compass-wrap{flex-wrap:wrap;justify-content:flex-start;gap:18px}.compass{width:210px!important;height:210px!important;min-width:210px;min-height:210px;flex-basis:210px}}@media(max-width:700px){.gsa-grid{grid-template-columns:1fr}}.compass-field{grid-column:1/3;min-width:0}.compass{width:220px!important;height:220px!important;min-width:220px;min-height:220px;max-width:none!important;max-height:none!important;aspect-ratio:1/1;flex:0 0 220px;border-radius:50%;box-sizing:border-box}.pv-chart-box{grid-column:3;grid-row:auto;align-self:center;display:flex;flex-direction:column;justify-content:center;min-width:0;width:100%;min-height:300px;margin:0!important;padding:14px 12px;overflow:hidden;border:1px solid #dce5e1;border-radius:10px;background:#fff}.pv-chart-box h4{font-size:13px;line-height:1.3;margin:0 0 8px;color:#245441}.pv-chart-box #pvChartCanvas{width:100%;overflow:hidden;flex:1;display:flex;align-items:center;justify-content:center}.pv-chart-box svg{display:block;width:100%;height:auto;min-width:0!important}.compass-wrap{display:flex;align-items:center;justify-content:center;gap:26px;flex-wrap:nowrap;padding:8px 4px}.compass{width:220px;height:220px;flex:0 0 220px}.compass-info{max-width:285px}.needle{height:68px}.needle i{border-bottom-width:60px}.needle b{top:60px}.compass-info{min-width:150px;flex:1}.compass-info .row{gap:6px}.compass-info .btn{font-size:12px;padding:8px 10px}.compass{width:250px;height:250px;max-width:75vw;max-height:75vw;border:3px solid #9aafa5;border-radius:50%;position:relative;background:radial-gradient(circle,#fff 0 57%,#edf4f0 58% 100%);box-shadow:inset 0 0 0 7px #f7faf8,0 2px 8px #00000012;touch-action:none;user-select:none;cursor:crosshair}.compass:before,.compass:after{content:'';position:absolute;background:#d7e2dc;left:50%;top:10%;width:1px;height:80%;transform:translateX(-50%)}.compass:after{transform:translateX(-50%) rotate(90deg)}.dir{position:absolute;font-size:11px;font-weight:bold;color:#52665d;z-index:2}.north{top:7px;left:50%;transform:translateX(-50%);color:#c0392b}.south{bottom:7px;left:50%;transform:translateX(-50%)}.east{right:8px;top:50%;transform:translateY(-50%)}.west{left:8px;top:50%;transform:translateY(-50%)}.ne{right:34px;top:32px}.se{right:34px;bottom:32px}.sw{left:34px;bottom:32px}.nw{left:34px;top:32px}.needle{position:absolute;z-index:3;left:50%;top:50%;width:8px;height:88px;transform:translate(-50%,-100%) rotate(180deg);transform-origin:50% 100%;pointer-events:none}.needle i{position:absolute;left:0;top:0;width:0;height:0;border-left:4px solid transparent;border-right:4px solid transparent;border-bottom:78px solid #e04435;filter:drop-shadow(0 1px 1px #777)}.needle b{position:absolute;left:1px;top:78px;width:6px;height:10px;background:#667b71;border-radius:0 0 3px 3px}.compass-center{position:absolute;z-index:4;left:50%;top:50%;width:15px;height:15px;border-radius:50%;background:#fff;border:4px solid #315c49;transform:translate(-50%,-50%);pointer-events:none}.compass-info{min-width:230px;flex:1}.compass-info .row{margin:12px 0}@media(max-width:650px){.grid{grid-template-columns:1fr 1fr}}@media(max-width:450px){.grid{grid-template-columns:1fr}}</style></head><body><header><b>☀ SOLAR QUOTE AUTOMATION</b> <button class='btn alt' type='button' onclick="location.href='/logout-now'" style='float:right'>Đăng xuất</button></header><main>
<section id="dashboardView"><div class="dash-head"><div><h1>Thư mục dự án</h1><p class="muted">Quản lý hồ sơ theo từng ngày, giống thư mục trên máy tính.</p></div></div>
<div class="panel folder-create"><h3>＋ Tạo thư mục mới</h3><div class="row folder-create-row"><div class="field"><label>Tên thư mục</label><input id="newFolderName" placeholder="Ví dụ: Khách hàng Hồng An"></div><div class="field"><label>Ngày tạo</label><input id="newFolderDate" type="date"></div><button class="btn" id="createFolderBtn" type="button">Tạo thư mục</button></div></div>
<div id="folderGroups"></div><div id="folderStatus" class="muted"></div></section>
<section id="folderView" style="display:none"><div class="row editor-toolbar"><button class="btn alt" type="button" onclick="backToFolders()">← Tất cả thư mục</button><div><b id="folderViewTitle">📁 Thư mục</b><div class="muted" id="folderViewDate"></div></div></div><div class="panel"><div class="row" style="justify-content:space-between;align-items:center"><div><h3 style="margin:0">Dự án trong thư mục</h3><p class="muted">Mở dự án đã lưu hoặc bắt đầu tạo dự án mới.</p></div><button class="btn" type="button" onclick="startProject()">＋ Bắt đầu tạo dự án</button></div><div id="folderProjectList" style="margin-top:14px"></div></div></section>
<section id="editorView" style="display:none"><div class="row editor-toolbar"><button class="btn alt" type="button" onclick="backToFolder()">← Quay lại thư mục</button><div><b id="currentFolderTitle">Thư mục dự án</b><div class="muted" id="currentFolderDate"></div></div></div>
<h1>Trung tâm báo giá điện mặt trời</h1><p class='muted'>Website MVP · Dữ liệu lưu trên máy chủ</p>
<div class='panel' id='loadSizingPanel' style='border:2px solid #cfe3d8'>
<div class='row' style='justify-content:space-between;align-items:center;gap:12px'><div><h3 style='margin:0'>⚡ Phân tích hóa đơn điện & phụ tải</h3><p class='muted' style='margin:5px 0 0'>Đọc hóa đơn EVN, tính giá điện và ước tính biểu đồ phụ tải nhà máy.</p></div><button class='btn alt' id='loadSizingToggle' type='button' onclick='toggleLoadSizing()' aria-expanded='false'>＋ Mở phần tính toán</button></div>
<div id='loadSizingContent' style='display:none;margin-top:14px'>
<p class='muted'>Chỉ cần tải hóa đơn EVN và khai báo lịch vận hành. Website tự ước tính phụ tải 30 phút từ sản lượng theo khung giờ; không cần xuất dữ liệu từ công tơ.</p>
<div class='grid'>
<div class='field' style='grid-column:1/-1'><label>Hóa đơn điện nhiều tháng / nhiều kỳ (PDF, JPG, PNG, WEBP — tối đa 50 file/lần)</label><div class='row'><input id='billFiles' type='file' accept='.pdf,.jpg,.jpeg,.png,.webp' multiple style='max-width:520px'><button class='btn alt' type='button' onclick='readBills()'>Đọc & tổng hợp hóa đơn</button><span id='billReadStatus' class='muted'></span><div id='billReadProgress' style='width:100%;margin-top:8px' aria-live='polite'></div></div><div class='muted'>Có thể chọn đồng thời hóa đơn VAT và bảng kê chi tiết. Website sẽ phát hiện cùng kỳ để tránh cộng trùng.</div></div><div class='field' style='grid-column:1/-1'><div id='billBatchSummary' class='muted'>Chưa có dữ liệu hóa đơn.</div><div id='billGroupSummary' style='margin-top:12px'></div><div class='field' style='margin-top:12px'><label>Nhóm hóa đơn dùng để tính sizing PV / pin</label><select id='billGroupSelect' onchange='applyBillGroup()' style='width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:8px'><option value=''>Chọn nhóm hóa đơn</option></select><small class='muted'>Chỉ cộng các hóa đơn cùng loại điện, cấp điện áp và hình thức giá; không trộn sinh hoạt với sản xuất/kinh doanh.</small></div><div style='overflow-x:auto;margin-top:8px'><table id='billBatchTable' style='width:100%;border-collapse:collapse;font-size:13px;display:none'><thead><tr><th>Tệp / kỳ hóa đơn</th><th>Loại điện</th><th>Hình thức giá</th><th>Sản lượng bậc</th><th>Cấp áp</th><th>Bình thường</th><th>Cao điểm</th><th>Thấp điểm</th><th>Tổng kWh</th><th>Sản lượng TB (kWh/ngày)</th><th>Tiền điện</th><th>Giá điện TB (đ/kWh)</th><th>Trạng thái</th></tr></thead><tbody id='billBatchRows'></tbody></table></div></div>
<div class='field'><label>Tổng điện năng hóa đơn (kWh)</label><input id='billMonthKwh' type='number' min='0' step='.01' placeholder='Nhập hoặc kiểm tra số đọc được'></div>
<div class='field'><label>Số ngày của kỳ hóa đơn</label><input id='billDays' type='number' min='1' step='1' value='30'></div>
<div class='field'><label>Sản lượng trung bình (kWh/ngày)</label><input id='billDailyAverage' type='number' readonly placeholder='Tự tính từ tổng kWh / số ngày'></div>
<div class='field'><label>Sản lượng trung bình tháng (kWh/tháng)</label><input id='billMonthlyAverageKwh' type='text' readonly placeholder='Quy đổi theo 30 ngày'></div>
<div class='field'><label>Tiền điện trung bình tháng (đồng/tháng)</label><input id='billMonthlyAverageMoney' type='text' readonly placeholder='Quy đổi theo 30 ngày'></div>
<div class='field'><label>Giá điện trung bình (đ/kWh)</label><input id='billAveragePrice' type='number' readonly placeholder='Tự tính từ tổng tiền / tổng kWh'></div>
<div style='grid-column:1/-1;border:1px solid #cbded4;border-radius:10px;padding:14px;background:#fbfdfc;margin:4px 0'>
<h3 style='margin:0 0 6px'>Quy đổi tiền điện sinh hoạt thành sản lượng theo 5 bậc</h3>
<p class='muted' style='margin:0 0 10px'>Nhập tổng tiền điện sinh hoạt. Công cụ tính ngược số kWh theo biểu giá 5 bậc áp dụng từ 01/11/2026 (đơn giá chưa VAT 8%).</p>
<div class='grid' style='grid-template-columns:repeat(2,minmax(0,1fr));gap:10px'>
<div class='field'><label>Tổng tiền điện (đồng)</label><input id='residentialBillAmount' type='number' min='0' step='1000' placeholder='Ví dụ: 1500000'></div>
<div class='field'><label style='display:flex;align-items:center;gap:8px'><input id='residentialVatIncluded' type='checkbox' checked style='width:auto'> Số tiền đã bao gồm VAT 8%</label><small class='muted'>Bỏ chọn nếu số tiền nhập chưa VAT.</small></div>
</div>
<button class='btn' type='button' onclick='calculateResidentialBill()'>Tính số điện theo 5 bậc</button>
<div id='residentialBillResult' class='muted' style='margin-top:10px'>Chưa tính tiền điện sinh hoạt.</div>
</div>
<div class='field'><label>Tỷ lệ điện năng ban ngày 06:00–18:00 (%)</label><input id='billDayShare' type='number' min='0' max='100' step='.1' value='65'></div>
<div class='field'><label>Tỷ lệ điện năng ban đêm 18:00–06:00 (%)</label><input id='billNightShare' type='number' readonly value='35'></div>
<div class='field'><label>Tỷ lệ pin đáp ứng điện ban đêm (%)</label><input id='billNightCoverage' type='number' min='0' max='100' step='1' value='100'><small class='muted'>Mặc định 100%; có thể giảm xuống 70–90% tùy mục tiêu đầu tư.</small></div>
<div class='field'><label>Số giờ nắng đỉnh PSH (h/ngày)</label><input id='billPsh' type='number' min='.5' max='8' step='.1' value='4'></div>
<div class='field'><label>Độ sâu xả pin DoD (%)</label><input id='billDod' type='number' min='50' max='100' step='1' value='90'></div>
<div class='field'><label>Hiệu suất sạc pin (%)</label><input id='billChargeEff' type='number' min='50' max='100' step='1' value='100'></div>
<div class='field'><label>Hiệu suất xả pin (%)</label><input id='billDischargeEff' type='number' min='50' max='100' step='1' value='100'></div>
</div>
<div class='row'><button class='btn' type='button' onclick='calculateLoadSizing()'>Tính điện năng & đề xuất</button><button class='btn alt' type='button' onclick='applyLoadSizing()'>Áp dụng vào cấu hình dự án</button></div>
<div id='loadSizingStatus' class='muted'>Chưa tính toán.</div>
<div id='loadSizingResults' style='display:none;margin-top:14px'>
<div class='grid'>
<div class='gsa-card'><div class='muted'>Tổng điện năng trung bình</div><h2 id='resDaily' style='margin:7px 0'>—</h2><div class='muted'>kWh/ngày</div></div>
<div class='gsa-card'><div class='muted'>Điện năng ban ngày</div><h2 id='resDay' style='margin:7px 0'>—</h2><div class='muted'>kWh/ngày</div></div>
<div class='gsa-card'><div class='muted'>Điện năng ban đêm</div><h2 id='resNight' style='margin:7px 0'>—</h2><div class='muted'>kWh/ngày</div></div>
<div class='gsa-card'><div class='muted'>Công suất PV đề xuất</div><h2 id='resPv' style='margin:7px 0;color:#16845b'>—</h2><div class='muted'>kWp DC · công suất mục tiêu</div><div id='resPvModules' style='font-size:12px;margin-top:8px;line-height:1.6'></div></div>
<div class='gsa-card'><div class='muted'>Dung lượng pin đề xuất</div><h2 id='resBattery' style='margin:7px 0;color:#16845b'>—</h2><div class='muted'>kWh danh định</div></div>
<div class='gsa-card'><div class='muted'>Điện năng cần pin cấp ban đêm</div><h2 id='resNightSupply' style='margin:7px 0'>—</h2><div class='muted'>kWh/ngày</div></div>
</div>
<p class='muted' id='loadSizingAssumptions' style='margin-top:10px'></p>
<details id='targetBatteryOptions' style='border:1px solid #cbded4;border-radius:10px;margin-top:14px;background:#fbfdfc'>
<summary style='cursor:pointer;list-style:none;padding:14px;font-weight:700;color:#176b4b;display:flex;align-items:center;justify-content:space-between'>
<span><span style='margin-right:8px'>＋</span>Tùy chọn dung lượng pin mục tiêu (mở rộng)</span><span class='muted' style='font-weight:400;font-size:12px'>Bấm để thu gọn / mở rộng</span>
</summary>
<div style='padding:0 14px 14px'>
<p class='muted' style='margin:0 0 10px'>Nhập điện năng pin cần cấp cho tải mỗi ngày (kWh). Website tự quy đổi ngược qua hiệu suất xả, hiệu suất sạc và DoD để tìm công suất PV.</p>
<div class='grid' style='grid-template-columns:repeat(2,minmax(0,1fr));gap:10px'>
<div class='field'><label>Điện năng muốn nạp vào pin mỗi ngày (kWh)</label><input id='targetBatteryDailyKwh' type='number' min='0' step='0.1' placeholder='Ví dụ: 100'></div>
<div class='field'><label>PV cấp tải ban ngày (%)</label><input id='targetPvDayCoverage' type='number' min='0' max='100' step='1' value='80'></div>
</div>
<div class='row' style='margin-top:8px'><button class='btn' type='button' onclick='calculateTargetBatteryPv()'>Tính PV theo pin mục tiêu</button><button class='btn alt' type='button' onclick='applyTargetBatteryPv()'>Áp dụng phương án này</button></div>
<div id='targetBatteryPvResult' class='muted' style='margin-top:10px'>Chưa tính phương án pin mục tiêu.</div>
</div>
</details>
<div style='border-top:1px solid #e2eae5;margin-top:16px;padding-top:14px'>
<h3 style='margin:0 0 6px'>Biểu đồ phụ tải (nhà máy 06:00–17:00 · hộ dân 06:00–20:00)</h3>
<p class='muted'>Chọn loại phụ tải để website áp dụng mô hình phù hợp: hộ dân dùng điện 24/24 với tỷ lệ ngày/đêm; nhà máy dùng lịch ca, ngày hoạt động và nghỉ trưa.</p>
<div class='field' style='max-width:520px;margin:10px 0'><label>Loại phụ tải</label><select id='loadCustomerMode' onchange='toggleLoadCustomerMode()'><option value='household'>1. Hộ dân</option><option value='factory' selected>2. Nhà máy</option></select></div>
<div id='householdLoadInfo' class='muted' style='display:none;margin:8px 0 12px'>Hộ dân sử dụng điện 24/24. Phân bổ điện năng theo hai khoảng bằng nhau: ban ngày 06:00–18:00 (12 giờ) và ban đêm 18:00–06:00 hôm sau (12 giờ). Biểu đồ vẫn hiển thị 06:00–20:00; phần 00:00–06:00 được tính vào điện năng ban đêm và nhu cầu pin nhưng không hiển thị. Không áp dụng lịch ca hoặc nghỉ trưa.<div style='margin-top:8px'><button class='btn' type='button' onclick='buildLoadProfileFromTou()'>Tạo phụ tải hộ dân từ kWh hóa đơn</button></div></div>
<div id='factoryLoadFields'>
<div class='grid' style='grid-template-columns:repeat(2,minmax(0,1fr));margin:10px 0'>
<div class='field'><label>Khung giờ tính toán</label><select id='touScheduleMode'><option value='old'>KHUNG GIỜ CŨ</option><option value='new'>KHUNG GIỜ MỚI</option></select></div>
<div class='field' style='align-self:end'><button class='btn' type='button' onclick='buildLoadProfileFromTou()'>Tạo đường phụ tải từ kWh hóa đơn</button></div>
</div>
<div class='muted' id='touScheduleDescription' style='margin:6px 0 12px'></div>
<h4 style='margin:12px 0 8px'>Chế độ phụ tải nhà máy</h4>
<div class='field' style='max-width:520px'><label>Kiểu vận hành</label><select id='plantLoadMode' onchange='togglePlantScheduleMode()'><option value='shift' selected>Theo lịch ca / có tải thay đổi theo giờ</option><option value='continuous'>Chạy liên tục 24/7, phụ tải ngày đêm như nhau</option></select></div>
<div id='plantScheduleFields'>
<h4 style='margin:12px 0 8px'>Lịch vận hành nhà máy</h4>
<div class='grid' style='grid-template-columns:repeat(3,minmax(0,1fr));gap:10px'>
<div class='field'><label>Số ngày hoạt động/tuần</label><select id='plantWorkDays'><option value='5'>5 ngày (Thứ Hai–Thứ Sáu)</option><option value='6' selected>6 ngày (Thứ Hai–Thứ Bảy)</option><option value='7'>7 ngày</option></select></div>
<div class='field'><label>Giờ bắt đầu ca</label><input id='plantStart' type='time' value='08:00'></div>
<div class='field'><label>Giờ kết thúc ca</label><input id='plantEnd' type='time' value='16:30'></div>
<div class='field'><label style='display:flex;align-items:center;gap:8px'><input id='plantLunchEnabled' type='checkbox' style='width:auto'> Có nghỉ trưa / giảm tải trong giờ nghỉ</label><small class='muted'>Bỏ chọn nếu nhà máy hoặc tải vận hành liên tục 24/24.</small></div>
<div class='field'><label>Bắt đầu nghỉ trưa</label><input id='plantLunchStart' type='time' value='11:00'></div>
<div class='field'><label>Kết thúc nghỉ trưa</label><input id='plantLunchEnd' type='time' value='13:00'></div>
<div class='field'><label>% máy dừng trong giờ nghỉ trưa</label><input id='plantLunchOff' type='number' min='0' max='100' value='50'></div>
<div class='field'><label>Mức phụ tải duy trì ngoài ca / ngày nghỉ (%)</label><input id='plantOutsideShare' type='number' min='0' max='100' value='10'></div>
</div>
<p class='muted'>Mặc định tải nền ngoài ca/ngày nghỉ là 10% để tạo ước tính ban đầu; hãy chỉnh theo thực tế từng nhà máy. Tỷ lệ máy dừng 80% nghĩa là còn khoảng 20% mức tải vận hành trong giờ nghỉ trưa.</p>
</div>
</div>
<p class='muted'>Website tự tạo 23 điểm phụ tải sau khi chọn nhóm hóa đơn; có thể chỉnh sửa theo thực tế. Nếu chỉ có hóa đơn một giá hoặc thiếu 3 khung giờ, không đủ cơ sở để tự dựng đường phụ tải theo TOU.</p>
<div class='row'><button class='btn alt' type='button' onclick='fillFlatLoadProfile()'>Phân bổ đều theo điện ban ngày</button><button class='btn alt' type='button' onclick='clearLoadProfile()'>Xóa dữ liệu nhập</button><button class='btn' type='button' onclick='drawLoadPvComparison()'>Vẽ / cập nhật biểu đồ</button></div>
<div id='loadProfileInputs' style='display:grid;grid-template-columns:repeat(6,minmax(90px,1fr));gap:8px'></div>
<div id='loadPvChart' style='display:none;margin-top:16px;border:1px solid #e2eae5;border-radius:9px;padding:10px'>
<div id='loadPvChartSvg'></div>
<div style='display:flex;gap:16px;flex-wrap:wrap;font-size:12px;margin-top:8px'><span style='color:#e53935'>━ Phụ tải (kW)</span><span style='color:#168a55'>▰ Solar (kW)</span><span id='batteryLegend' style='color:#2878d0'>▰ Miền công suất xả pin (kW)</span></div>
<div id='loadPvMetrics' class='grid' style='margin-top:12px'></div>
<p class='muted' id='loadPvNote' style='margin:10px 0 0'></p>
</div>
</div>
</div>
</div></div>
<details class='panel' id='pvwattsOptions' style='border:2px solid #cfe3d8'>
<summary style='cursor:pointer;list-style:none;display:flex;align-items:center;justify-content:space-between;gap:12px'>
<span><span style='font-weight:700;color:#176b4b;font-size:15px'>☀️ Mô phỏng sản lượng PVWatts V8</span><span class='muted' style='display:block;margin-top:5px'>Nhập tọa độ và thông số để mô phỏng sản lượng điện mặt trời.</span></span>
<span class='btn alt' id='pvwattsToggleLabel' style='white-space:nowrap;pointer-events:none'>＋ Mở mô phỏng</span>
</summary>
<div style='padding-top:12px'><p class='muted'>Nhập tọa độ và thông số mô phỏng, sau đó bấm lấy dữ liệu trực tiếp từ PVWatts.</p><div class='grid'>
<div class='field' style='grid-column:1/-1'><label>Tọa độ công trình (dán Latitude, Longitude)</label><input id='coordPaste' type='text' placeholder='Ví dụ: 10.821006371676384, 106.68872440336703' autocomplete='off'><div id='coordStatus' class='muted'>Dán một dòng tọa độ Google Maps; website tự tách vĩ độ và kinh độ.</div><input id='lat' type='hidden'><input id='lon' type='hidden'></div>
<div class='field'><label>DC System Size (kW)</label><input id='pvSystemSize' type='number' step='.01' placeholder='Tự lấy từ công suất PV nếu để trống'></div>
<div class='field'><label>Module Type</label><select id='moduleType'><option value='0'>Standard</option><option value='1'>Premium</option><option value='2'>Thin film</option></select></div>
<div class='field'><label>Array Type</label><select id='arrayType'><option value='0' selected>Fixed - Open Rack</option><option value='1'>Fixed - Roof Mounted</option><option value='2'>1-Axis</option><option value='3'>1-Axis Backtracking</option><option value='4'>2-Axis</option></select></div>
<div class='field'><label>Tilt (độ)</label><input id='tilt' type='number' value='20' step='.1'></div>
<div class='field compass-field'><label>Hướng mặt tấm pin (Azimuth)</label><div class='compass-wrap'><div class='compass' id='compass' aria-label='La bàn chọn hướng PV'><span class='dir north'>BẮC</span><span class='dir east'>ĐÔNG</span><span class='dir south'>NAM</span><span class='dir west'>TÂY</span><span class='dir ne'>ĐB</span><span class='dir se'>ĐN</span><span class='dir sw'>TN</span><span class='dir nw'>TB</span><div class='needle' id='compassNeedle'><i></i><b></b></div><div class='compass-center'></div></div><div class='compass-info'><div><b id='azimuthText'>180° · Hướng Nam</b><div class='muted'>Xoay kim theo hướng mặt tấm pin thực tế đang quay.</div></div><div class='row'><button type='button' class='btn' id='confirmAzimuth' onclick='confirmAzimuth()'>Xác nhận hướng</button><button type='button' class='btn alt' id='editAzimuth' onclick='editAzimuth()'>Chỉnh sửa</button></div><div class='field' id='manualAzimuthWrap' style='display:none;max-width:180px'><label>Nhập góc thủ công (0–359,9°)</label><input id='azimuth' type='number' value='180' min='0' max='359.9' step='.1'></div><div class='muted' id='azimuthStatus'>Hướng hiện tại đã xác nhận.</div></div></div></div>
<div id='pvChart' class='pv-chart-box' style='display:none'><h4>Biểu đồ sản lượng theo tháng</h4><div id='pvChartCanvas'></div></div>
<div class='field'><label>System Losses (%)</label><input id='losses' type='number' value='14.08' step='.01'></div>
<div class='field'><label>DC to AC Size Ratio</label><input id='dcAcRatio' type='number' value='1.25' step='.01'></div>
<div class='field'><label>Inverter Efficiency (%)</label><input id='invEff' type='number' value='97' step='.1'></div>
<div class='field'><label>Ground Coverage Ratio</label><input id='gcr' type='number' value='.4' step='.01'></div>
<div class='field'><label>Albedo</label><select id='albedoMode'><option value='weather' selected>Theo dữ liệu thời tiết (Weather file)</option><option value='manual'>Nhập thủ công</option></select></div>
<div class='field' id='albedoField' style='display:none'><label>Albedo (0–1)</label><input id='albedo' type='number' value='.2' step='.01'></div>
</div><div class='row'><button class='btn' onclick='runPVWatts()'>Tính toán PVWatts</button><span id='pvStatus' class='muted'></span></div><div id='pvResult' class='muted'>Chưa có kết quả mô phỏng.</div></div>
</div></details>
<div class='panel'><h3>Khách hàng & công trình</h3><div class='grid'><div class='field'><label>Tên dự án / Khách hàng / Công ty</label><input id='projectName' placeholder='Tên dự án hoặc tên khách hàng / công ty'></div><div class='field'><label>Số điện thoại</label><input id='phone' placeholder='Tự lấy từ báo giá nếu có; nếu không có, vui lòng nhập'></div><div class='field'><label>Địa điểm công trình</label><input id='address' placeholder='Tự lấy từ báo giá PDF nếu có'></div><div class='field'><label>Ngày báo giá</label><input id='date' type='date'></div><div class='field'><label>Ngày hiệu lực báo giá (để trống = +7 ngày)</label><input id='quoteValidUntil' type='date'></div><div class='field'><label>Ngành nghề</label><input id='industry' placeholder='Ví dụ: Sản xuất bao bì'></div><div class='field'><label>Diện tích mái (m²)</label><input id='roof' type='number'></div><div class='field'><label>Tổng mức đầu tư trước VAT (VNĐ)</label><input id='investmentSubtotal' type='number' min='0' step='1' placeholder='Tự đọc từ báo giá PDF'></div><div class='field'><label>Tiền VAT (VNĐ)</label><input id='investmentVat' type='number' min='0' step='1' placeholder='Tự đọc từ báo giá PDF'></div><div class='field'><label>Tổng mức đầu tư sau VAT (VNĐ)</label><input id='investmentTotal' type='number' min='0' step='1' placeholder='Tự đọc từ báo giá PDF hoặc nhập thủ công'></div><div class='field'><label>Bức xạ (kWh/m²/năm)</label><input id='solarRadiation' type='number' step='.1'></div><div class='field'><label>Kết cấu mái</label><input id='roofType' placeholder='Ví dụ: Kèo thép, xà gồ C'></div><div class='field'><label>Điều kiện mái</label><input id='roofCondition' placeholder='Ví dụ: Mái mới, không dột'></div><div class='field'><label>Thói quen dùng điện</label><textarea id='usagePattern' rows='2' style='width:100%;padding:10px;border:1px solid #ccd9d2;border-radius:7px' placeholder='Ví dụ: Hoạt động 2 ca, từ 7h đến 22h; nghỉ Chủ nhật...'></textarea></div></div></div>
<div class='panel'><h3>Cấu hình hệ thống</h3><div class='grid'><div class='field'><label>Thương hiệu tấm pin</label><input id='panelBrand'></div><div class='field'><label>Công suất tấm pin (Wp)</label><input id='panelW' type='number' value='625'></div><div class='field'><label>Số lượng tấm</label><input id='panelQty' type='number' value='30'></div><div class='field'><label>Công suất PV (kWp)</label><input id='pv' readonly></div><div class='field'><label>Inverter / Model</label><input id='invModel'></div><div class='field'><label>Công suất inverter (kW)</label><input id='invKw' type='number'></div><div class='field'><label>Số pha</label><select id='phase'><option>1 pha</option><option>3 pha</option></select></div><div class='field'><label>Pin lưu trữ / Model</label><input id='batModel'></div><div class='field'><label>Dung lượng mỗi bộ (kWh)</label><input id='batKwh' type='number' step='.01'></div><div class='field'><label>Số lượng bộ pin</label><input id='batQty' type='number' value='0'></div><div class='field'><label>Tổng dung lượng pin (kWh)</label><input id='batTotal' readonly></div></div></div>
<div class='panel'><h3>Điện năng & báo giá</h3><div class='grid'><div class='field'><label for='load' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>Tiêu thụ tháng (kWh)</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_load' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('load')"> Tự động</span></label><input id='load' type='number'  ></div><div class='field'><label for='annual' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>Sản lượng ĐMT năm (kWh)</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_annual' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('annual')"> Tự động</span></label><input id='annual' type='number'  ></div><div class='field'><label for='self'>Tự tiêu thụ dự kiến (%)</label><input id='self' type='number' min='0' max='100' step='1' value='100' ></div><div class='field'><label for='price' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>Giá điện trung bình có VAT (đ/kWh)</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_price' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('price')"> Tự động</span></label><input id='price' type='number'  ></div><div class='field'><label for='pre' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>Tổng trước VAT (VNĐ)</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_pre' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('pre')"> Tự động</span></label><input id='pre' type='number'  ></div><div class='field'><label for='vat' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>VAT (%)</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_vat' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('vat')"> Tự động</span></label><input id='vat' type='number' value='8' ></div><div class='field'><label for='vatMoney' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>Tiền VAT</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_vatMoney' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('vatMoney')"> Tự động</span></label><input id='vatMoney' type='number'  readonly></div><div class='field'><label for='grand' style='display:flex;align-items:center;justify-content:space-between;gap:8px'><span>Tổng sau VAT</span><span style='display:inline-flex;align-items:center;gap:4px;font-weight:400;font-size:11px;color:#168a55'><input id='auto_grand' type='checkbox' checked style='width:auto;margin:0' onchange="toggleQuoteAuto('grand')"> Tự động</span></label><input id='grand' type='number'  readonly></div><div class='field'><label for='inverterPrice'>Tổng giá Inverter (VNĐ, gồm VAT)</label><input id='inverterPrice' type='number' min='0' step='1' placeholder='Nhập giá inverter'></div><div class='field'><label for='batteryPrice'>Tổng giá pin lưu trữ (VNĐ, gồm VAT)</label><input id='batteryPrice' type='number' min='0' step='1' placeholder='Nhập giá pin lưu trữ'></div></div></div>
<div class='panel' id='financialPanel'><div class='row' style='justify-content:space-between;align-items:center'><div><h3 style='margin:0'>Phân tích tài chính</h3><p class='muted'>Sử dụng mặc định theo file mẫu. Chỉ điều chỉnh khi cần thiết.</p></div><button class='btn alt' id='financialToggle' type='button' onclick='toggleFinancialAdvanced()'>⚙ Điều chỉnh chuyên sâu PTTC</button></div><div id='financialAdvanced' style='display:none;margin-top:16px'><p class='muted'>Ô để trống sẽ giữ nguyên thông số trong file Excel gốc. Chỉ các ô có nhập giá trị mới được thay đổi.</p><div class='grid'><div class='field'><label>Tỷ lệ lạm phát (%/năm)</label><input id='finInflation' type='number' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Mức tăng giá điện (%/năm)</label><input id='finElectricityEscalation' type='number' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Tỷ lệ tiêu thụ điện (%)</label><input id='finSelfConsumption' type='number' step='0.01' placeholder='Giữ nguyên theo thông số dự án'></div><div class='field'><label>Tỷ lệ vay vốn (%)</label><input id='finLoanRatio' type='number' min='0' max='100' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Thời hạn vay (năm)</label><input id='finLoanTerm' type='number' min='0' step='1' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Lãi suất vay (%/năm)</label><input id='finLoanRate' type='number' min='0' step='0.01' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Thời gian khấu hao (năm)</label><input id='finDepreciation' type='number' min='1' step='1' placeholder='Giữ nguyên mặc định'></div><div class='field'><label>Chi phí vận hành (triệu đồng/năm)</label><input id='finOperatingCost' type='number' min='0' step='0.01' placeholder='Giữ nguyên mặc định'></div></div><div class='row'><button class='btn' type='button' onclick='applyFinancialOverrides()'>Thay đổi & tạo báo cáo</button><button class='btn alt' type='button' onclick='resetFinancialOverrides()'>Xóa điều chỉnh</button></div><div id='financialStatus' class='muted' aria-live='polite'></div></div><div class='row' style='align-items:center'><button class='btn alt' type='button' onclick="document.getElementById('projectTemplateFolder').click()">Chọn thư mục mẫu của dự án</button><input id='projectTemplateFolder' type='file' webkitdirectory directory multiple style='display:none' onchange='uploadProjectTemplates()'><span id='projectTemplateFolderName' class='muted'>Chưa chọn thư mục mẫu</span><button class='btn' type='button' onclick='generateFinancialReport()'>Tạo báo cáo PTTC</button></div><div id='projectTemplateStatus' class='muted' aria-live='polite'></div><div id='financialReportStatus' class='muted' aria-live='polite'></div></div>
<div class='panel' id='presentationPanel'><h3>Xuất báo giá PowerPoint</h3><p class='muted'>Chọn thư mục chứa PTTC khung (.xlsx) và báo giá PowerPoint khung (.pptx) cho dự án hiện tại. Mỗi dự án chọn thư mục riêng; website kiểm tra đủ file trước khi cho phép xuất.</p><div class='row' style='align-items:center'><button class='btn alt' type='button' onclick="document.getElementById('projectTemplateFolder').click()">Chọn thư mục mẫu</button><button class='btn' type='button' onclick='generatePresentationReport()'>Xuất báo giá PowerPoint</button></div><div id='presentationStatus' class='muted' aria-live='polite'></div></div>
<div class='panel'><h3>Ảnh phân tích tài nguyên mặt trời – Global Solar Atlas</h3><p class='muted'>Nhập tọa độ ở phần PVWatts, lưu dự án, sau đó hệ thống tự mở Global Solar Atlas và chụp 4 khu vực dữ liệu thực tế.</p><div class='row'><button class='btn' id='gsaCaptureBtn' onclick='captureGlobalSolarAtlas()'>Tự động lấy 4 hình ảnh</button><span id='gsaStatus' class='muted'>Chưa chụp ảnh.</span></div><div id='gsaResult' style='display:none'><div class='gsa-grid'><section class='gsa-card'><h4>1. Thông số bức xạ & PV</h4><img id='gsaImage1' alt='Thông số bức xạ và PV' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section><section class='gsa-card'><h4>2. Bản đồ vị trí công trình</h4><img id='gsaImage2' alt='Bản đồ vị trí công trình' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section><section class='gsa-card'><h4>3. Horizon & Sunpath</h4><img id='gsaImage3' alt='Horizon và Sunpath' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section><section class='gsa-card'><h4>4. Bản đồ tiềm năng PVOUT</h4><img id='gsaImage4' alt='Bản đồ PVOUT' style='width:100%;height:auto;border:1px solid #dce5e1;border-radius:6px'></section></div><p class='muted'>Ảnh chụp trực tiếp từ Global Solar Atlas. Nguồn: © World Bank Group, ESMAP và Solargis.</p></div></div>
<div class='panel'><h3>Hình ảnh phục vụ báo giá PowerPoint</h3><p class='muted'>Tải riêng ảnh hiện trạng mái và ảnh bố trí tấm PV. Ảnh được lưu theo đúng dự án đang mở.</p><div class='grid'><div class='field'><label>1. Hình ảnh mái thực tế</label><input id='roofRealImages' type='file' accept='image/jpeg,image/png,image/webp' multiple><div class='row'><button class='btn' type='button' onclick="uploadProjectImages('roof-real')">Tải ảnh mái thực tế</button></div><div id='roofRealImageList' class='muted'>Chưa có ảnh.</div></div><div class='field'><label>2. Hình ảnh mái bố trí PV</label><input id='roofPvImages' type='file' accept='image/jpeg,image/png,image/webp' multiple><div class='row'><button class='btn' type='button' onclick="uploadProjectImages('roof-pv')">Tải ảnh bố trí PV</button></div><div id='roofPvImageList' class='muted'>Chưa có ảnh.</div></div></div></div><div class='panel'><h3>Đọc báo giá PDF – tự điền tổng mức đầu tư</h3><p class='muted'>Tải báo giá PDF của nhà cung cấp. Website tự đọc tổng trước VAT, tiền VAT và tổng sau VAT, sau đó điền riêng vào ba ô trong phần Khách hàng & công trình.</p><div class='row' style='align-items:center'><input id='quotePdfFile' type='file' accept='.pdf,application/pdf' style='max-width:360px'><button class='btn' type='button' onclick='parseQuotePdf()'>Đọc báo giá & điền tự động</button></div><div id='quotePdfStatus' class='muted' aria-live='polite'></div></div><div class='panel'><h3>Hồ sơ đầu vào</h3><input id='files' type='file' multiple><div class='row'><button class='btn' onclick='upload()'>Tải file lên máy chủ</button></div><div id='fileList' class='muted'></div></div>
<div class='panel'><h3>Dự án đã lưu</h3><div class='row'><button class='btn' onclick='save()'>Lưu dự án</button><button class='btn alt' onclick='fresh()'>Dự án mới</button><button class='btn alt' onclick='exportJson()'>Xuất JSON</button></div><div id='projects'>Đang tải...</div><div id='status' class='muted'></div></div></section></main>
<script>
/* Runtime diagnostics: surface initialization errors instead of leaving a blank folder page. */
(function(){
 function showAppError(label,detail){
  let box=document.getElementById('appRuntimeError');
  if(!box){box=document.createElement('div');box.id='appRuntimeError';box.style.cssText='position:fixed;z-index:20000;left:12px;right:12px;bottom:12px;padding:14px 16px;background:#fff0f0;border:2px solid #c62828;border-radius:8px;color:#7f1111;font:14px/1.45 Arial,sans-serif;white-space:pre-wrap;box-shadow:0 4px 20px #0003';document.body.appendChild(box)}
  box.textContent='LỖI KHỞI TẠO WEBSITE ('+label+'):\\n'+String(detail||'Không rõ lỗi');
 }
 window.addEventListener('error',function(e){showAppError('JavaScript',e.message+'\\n'+(e.filename||'')+':'+(e.lineno||0)+':'+(e.colno||0));});
 window.onerror=function(message,source,line,column,error){showAppError('Lỗi cú pháp JavaScript',String(message)+'\\nVị trí: '+line+':'+column+'\\nNguồn: '+source+'\\n'+(error&&error.stack||''));return false;};
 window.addEventListener('unhandledrejection',function(e){const r=e.reason;showAppError('Promise',r&&(r.stack||r.message)||r);});
 window.addEventListener('DOMContentLoaded',function(){setTimeout(function(){const d=document.getElementById('newFolderDate');if(d&&!d.value){
 const ss=Array.from(document.scripts);
 const tail=ss[ss.length-1];
 let parse='không kiểm tra được';
 try{if(tail&&tail.textContent)new Function(tail.textContent),parse='cú pháp hợp lệ';else parse='script cuối rỗng'}
 catch(e){
 parse='LỖI CÚ PHÁP: '+e.message;
 // Ask the browser's native script parser for the exact source line/column.
 try{
  let captured=null;
  const capture=ev=>{if(ev&&ev.message)captured={message:ev.message,line:ev.lineno,column:ev.colno,source:ev.filename}};
  window.addEventListener('error',capture,true);
  const probe=document.createElement('script');
  probe.textContent='if(false){\\n'+tail.textContent+'\\n}';
  document.head.appendChild(probe);
  probe.remove();
  window.removeEventListener('error',capture,true);
  if(captured)parse+='; DÒNG/CỘT: '+captured.line+':'+captured.column+'; '+captured.message+'; '+captured.source;
  else parse+='; Trình duyệt chưa trả vị trí parser; STACK: '+String(e.stack||'').slice(0,500);
 }catch(probeError){parse+='; Không lấy được vị trí parser: '+probeError.message}
}
 const info='Bước: '+(window.__solarInitStage||'main script chưa chạy')+'; script tags='+ss.length+'; script cuối='+((tail&&tail.textContent||'').length)+' ký tự; '+parse+'; có marker='+(!!(tail&&tail.textContent.includes('__solarInitStage')));
 showAppError('Khởi tạo bị dừng',info);
}},1200);});
})();
</script>
<script>
window.__solarInitStage='main-script-entered';
const pvDetails=document.getElementById('pvwattsOptions');
if(pvDetails){pvDetails.addEventListener('toggle',()=>{const b=document.getElementById('pvwattsToggleLabel');if(b)b.textContent=pvDetails.open?'− Thu gọn mô phỏng':'＋ Mở mô phỏng'})}
let pid=null,pvwattsResult=null,gsaImages=[],financialReportId=null,currentFolderId=null,currentFolderName='',currentFolderDate='';const ids=['lat','lon','pvSystemSize','moduleType','arrayType','tilt','azimuth','losses','dcAcRatio','invEff','gcr','albedoMode','albedo','projectName','phone','address','date','quoteValidUntil','industry','solarRadiation','investmentSubtotal','investmentVat','investmentTotal','roofCondition','usagePattern','roof','roofType','plantWorkDays','plantStart','plantEnd','plantLunchStart','plantLunchEnd','plantLunchOff','plantOutsideShare','panelBrand','panelW','panelQty','invModel','invKw','phase','batModel','batKwh','batQty','load','annual','self','price','pre','vat','inverterPrice','batteryPrice'];const $=x=>document.getElementById(x),n=x=>Number($(x).value)||0;
function createFolderFromClick(event){if(event)event.preventDefault();const status=$('folderStatus');if(status)status.textContent='Đang gửi yêu cầu tạo thư mục...';createFolder().catch(err=>{if(status)status.textContent='Không tạo được thư mục: '+(err&&err.message?err.message:String(err));console.error('createFolderFromClick',err)})}
function setQuoteAuto(id,value){const cb=$('auto_'+id),el=$(id);if(cb&&cb.checked&&value!==null&&value!==undefined&&Number.isFinite(Number(value)))el.value=Number(value).toFixed(id==='load'||id==='annual'?2: id==='price'?2:0)}
function toggleQuoteAuto(id){const cb=$('auto_'+id),el=$(id);if(el)el.readOnly=!!cb.checked&&(id==='vatMoney'||id==='grand');syncQuoteAuto();calc()}
function syncQuoteAuto(){
 setQuoteAuto('load',n('billMonthKwh')||null);
 setQuoteAuto('annual',pvwattsResult&&pvwattsResult.annual!=null?pvwattsResult.annual:null);
 setQuoteAuto('price',n('billAveragePrice')||null);
 const kwh=n('billMonthKwh'),avg=n('billAveragePrice');
 setQuoteAuto('pre',kwh>0&&avg>0? kwh*avg/(1+n('vat')/100):null);
 const vatInput=$('vat');if(vatInput&&$('auto_vat')?.checked)vatInput.value=8;
 calc();
}
function calc(){let kw=n('panelW')*n('panelQty')/1000,v=n('pre')*n('vat')/100;$('pv').value=kw.toFixed(3);if(loadSizingResult&&$('resPvModules'))updatePvModuleRecommendation(loadSizingResult.pvKwp);$('batTotal').value=(n('batKwh')*n('batQty')).toFixed(2);if($('auto_vatMoney').checked)$('vatMoney').value=v.toFixed(0);if($('auto_grand').checked)$('grand').value=(n('pre')+v).toFixed(0)}
document.addEventListener('input',e=>{if(e.target.tagName==='INPUT'){if(['billMonthKwh','billAveragePrice','vat'].includes(e.target.id))syncQuoteAuto();else calc()}});
let financialOverrides={};
const financialFields={finInflation:'inflation',finElectricityEscalation:'electricityEscalation',finSelfConsumption:'selfConsumption',finLoanRatio:'loanRatio',finLoanTerm:'loanTerm',finLoanRate:'loanRate',finDepreciation:'depreciationYears',finOperatingCost:'operatingCost'};
function readFinancialOverrides(){let out={};for(const [id,key] of Object.entries(financialFields)){let raw=$(id).value.trim();if(raw!==''){let v=Number(raw);if(!Number.isFinite(v))throw Error('Giá trị '+key+' không hợp lệ.');out[key]=key==='selfConsumption'?v/100:v}}return out}
function showFinancialOverrides(){for(const [id,key] of Object.entries(financialFields))$(id).value=financialOverrides[key]===undefined?'':(key==='selfConsumption'?financialOverrides[key]*100:financialOverrides[key])}
function toggleFinancialAdvanced(){let el=$('financialAdvanced'),open=el.style.display==='none';el.style.display=open?'block':'none';$('financialToggle').textContent=open?'Ẩn điều chỉnh chuyên sâu':'⚙ Điều chỉnh chuyên sâu PTTC'}
async function applyFinancialOverrides(){try{financialOverrides=readFinancialOverrides();$('financialStatus').textContent='Đã áp dụng '+Object.keys(financialOverrides).length+' thông số điều chỉnh. Đang lưu theo dự án...';if(pid){let p=payload();p.financialOverrides=financialOverrides;await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});$('financialStatus').textContent='Đã lưu thông số PTTC. Đang tạo báo cáo theo các thay đổi...';await generateFinancialReport()}else{$('financialStatus').textContent='Hãy lưu dự án trước khi áp dụng điều chỉnh và tạo báo cáo PTTC.'}}catch(e){$('financialStatus').textContent=e.message}}
function resetFinancialOverrides(){financialOverrides={};showFinancialOverrides();$('financialStatus').textContent='Đã xóa các điều chỉnh; sẽ sử dụng thông số mặc định của mẫu PTTC.'}
async function uploadProjectTemplates(){if(!pid){$('projectTemplateStatus').textContent='Hãy lưu dự án trước khi chọn thư mục mẫu.';return}const files=Array.from($('projectTemplateFolder').files||[]);if(!files.length){$('projectTemplateStatus').textContent='Bạn chưa chọn thư mục hoặc thư mục đang trống.';return}const xlsx=files.filter(f=>f.name.toLowerCase().endsWith('.xlsx')),pptx=files.filter(f=>f.name.toLowerCase().endsWith('.pptx'));const errors=[];if(xlsx.length!==1)errors.push('Cần đúng 1 file PTTC khung .xlsx (hiện có '+xlsx.length+').');if(pptx.length!==1)errors.push('Cần đúng 1 file báo giá PowerPoint khung .pptx (hiện có '+pptx.length+').');if(errors.length){$('projectTemplateStatus').textContent='Thư mục chưa hợp lệ: '+errors.join(' ');return}const fd=new FormData();files.forEach(f=>fd.append('files',f,f.name));try{$('projectTemplateStatus').textContent='Đang kiểm tra cấu trúc Excel và PowerPoint...';const r=await fetch('/api/projects/'+pid+'/templates',{method:'POST',body:fd});let d={};try{d=await r.json()}catch(_){}if(!r.ok)throw Error(d.detail||'Không thể sử dụng thư mục mẫu');$('projectTemplateStatus').textContent='Thư mục hợp lệ: '+d.financial_name+' và '+d.presentation_name+' đã được gắn với dự án hiện tại.';$('projectTemplateFolderName').textContent='Đã chọn: '+(files[0].webkitRelativePath||files[0].name).split('/')[0];}catch(e){$('projectTemplateStatus').textContent='Lỗi thư mục mẫu: '+e.message}}
async function generateFinancialReport(){if(!pid){$('financialReportStatus').textContent='Hãy lưu dự án trước khi tạo báo cáo PTTC.';return}try{$('financialReportStatus').textContent='Đang tính toán và tạo Excel + PDF (trang 5–6)...';let p=payload();let r=await api('/api/projects/'+pid+'/financial-report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({payload:p})});financialReportId=r.excel_id||r.id;let savedPayload=payload();savedPayload.financialReportId=r.excel_id||r.id;savedPayload.financialPdfId=r.pdf_id;await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(savedPayload)});$('financialReportStatus').innerHTML='Đã tạo báo cáo PTTC: <a href="/api/files/'+r.excel_id+'" download>⬇ Tải Excel</a> &nbsp; | &nbsp; <a href="#" id="financialPdfPreviewLink">▣ Xem trước PDF (trang 5–6)</a>';$('financialPdfPreviewLink').addEventListener('click',ev=>{ev.preventDefault();openFinancialPdfPreview(r.pdf_id)});await listFiles()}catch(e){$('financialReportStatus').textContent='Không tạo được báo cáo PTTC: '+e.message}} 
function openFinancialPdfPreview(fid){const modal=$('financialPdfModal');$('financialPdfFrame').src='/api/files/'+fid+'/preview';$('financialPdfDownload').href='/api/files/'+fid;$('financialPdfModal').style.display='flex';document.body.style.overflow='hidden'}
function closeFinancialPdfPreview(){const modal=$('financialPdfModal');$('financialPdfFrame').src='about:blank';modal.style.display='none';document.body.style.overflow=''}
async function parseQuotePdf(){const f=$('quotePdfFile').files[0];if(!f){$('quotePdfStatus').textContent='Bạn hãy chọn báo giá PDF trước.';return}if(f.type!=='application/pdf'&&!f.name.toLowerCase().endsWith('.pdf')){$('quotePdfStatus').textContent='Chỉ hỗ trợ file PDF.';return}const fd=new FormData();fd.append('file',f);try{$('quotePdfStatus').textContent='Đang đọc nội dung báo giá PDF...';const resp=await fetch('/api/quote-parse',{method:'POST',body:fd});let data={};try{data=await resp.json()}catch(_){}if(!resp.ok)throw Error(data.detail||'Không đọc được báo giá.');const put=(id,v)=>{if(v!==null&&v!==undefined&&v!==''&&$(id))$(id).value=v};put('investmentSubtotal',data.subtotal);put('investmentVat',data.vat);put('investmentTotal',data.total_including_vat);put('projectName',data.customer_name);put('phone',data.customer_phone);put('address',data.site_address);put('panelBrand',data.panel_brand);put('panelW',data.panel_wp);put('panelQty',data.panel_qty);put('invModel',data.inverter_model);put('invKw',data.inverter_kw);put('phase',data.phase);put('batModel',data.battery_model);put('batKwh',data.battery_kwh);put('batQty',data.battery_qty);if(data.panel_wp&&data.panel_qty)put('pv',(Number(data.panel_wp)*Number(data.panel_qty)/1000).toFixed(3));if(data.battery_kwh&&data.battery_qty)put('batTotal',(Number(data.battery_kwh)*Number(data.battery_qty)).toFixed(2));calc();put('pre',data.subtotal);put('vatMoney',data.vat);put('grand',data.total_including_vat);const fields=[['Tổng trước VAT',data.subtotal],['Tiền VAT',data.vat],['Tổng sau VAT',data.total_including_vat],['Hãng pin',data.panel_brand],['Công suất tấm',data.panel_wp],['Số tấm',data.panel_qty],['Model inverter',data.inverter_model],['Công suất inverter',data.inverter_kw],['Số pha',data.phase],['Model pin lưu trữ',data.battery_model],['Dung lượng mỗi bộ',data.battery_kwh],['Số bộ pin',data.battery_qty]];const found=fields.filter(x=>x[1]!==null&&x[1]!==undefined&&x[1]!=='').map(x=>x[0]).join(', ');const missing=fields.filter(x=>x[1]===null||x[1]===undefined||x[1]==='').map(x=>x[0]).join(', ');$('quotePdfStatus').textContent='Đã đọc '+esc(f.name)+'. Tổng trước VAT: '+(data.subtotal==null?'chưa nhận diện':Number(data.subtotal).toLocaleString('vi-VN')+' đ')+'; VAT: '+(data.vat==null?'chưa nhận diện':Number(data.vat).toLocaleString('vi-VN')+' đ')+'; Tổng sau VAT: '+(data.total_including_vat==null?'chưa nhận diện':Number(data.total_including_vat).toLocaleString('vi-VN')+' đ')+'. Đã tự điền: '+(found||'chưa nhận diện được dữ liệu')+'. '+(missing?'Chưa nhận diện được: '+missing+'.':'Đã nhận diện đủ dữ liệu.')+' '+(data.method==='text'?'Đọc văn bản PDF':'Đọc OCR')+'.';}catch(e){$('quotePdfStatus').textContent='Không đọc được báo giá: '+e.message}}


async function generatePresentationReport(){if(!pid){$('presentationStatus').textContent='Hãy lưu dự án trước khi xuất PowerPoint.';return}try{$('presentationStatus').textContent='Đang tạo PowerPoint theo form gốc...';let r=await api('/api/projects/'+pid+'/presentation-report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({payload:payload()})});$('presentationStatus').innerHTML='Đã xuất báo giá: <a href="/api/files/'+r.id+'">'+esc(r.name)+'</a>';await listFiles()}catch(e){$('presentationStatus').textContent='Không xuất được PowerPoint: '+e.message}}


function payload(){let o={};ids.forEach(k=>o[k]=$(k).value);o.customer=o.projectName||'';o.contact='';o.loadSizingResult=loadSizingResult;o.loadProfileKw=loadProfileTimes.map((_,i)=>$('loadProfile'+i)?.value??'');o.touScheduleMode=$('touScheduleMode')?.value||'old';o.financialOverrides=financialOverrides;o.financialReportId=financialReportId;o.pvwattsResult=pvwattsResult;o.globalSolarAtlasResult=globalSolarAtlasResult;o.gsaImages=gsaImages;return o}function fill(o){ids.forEach(k=>{if(o[k]!==undefined&&$(k))$(k).value=o[k]});if(!o.projectName&&o.customer&&$('projectName'))$('projectName').value=o.customer;loadSizingResult=o.loadSizingResult||null;if(o.touScheduleMode&&$('touScheduleMode')){$('touScheduleMode').value=o.touScheduleMode;updateTouScheduleDescription()}if(Array.isArray(o.loadProfileKw)){o.loadProfileKw.forEach((v,i)=>{if(i<23&&$('loadProfile'+i))$('loadProfile'+i).value=v??''});if(o.loadProfileKw.length===23&&o.loadProfileKw.some(v=>v!==''&&v!==null&&v!==undefined)){try{drawLoadPvComparison()}catch(_){}}}if(loadSizingResult){$('billMonthKwh').value=loadSizingResult.monthlyKwh||'';$('billDays').value=loadSizingResult.billDays||30;$('billDayShare').value=loadSizingResult.dayShare||65;calcBillNightShare();calculateLoadSizing()}else{$('loadSizingResults').style.display='none'}financialOverrides=o.financialOverrides||{};financialReportId=o.financialReportId||null;showFinancialOverrides();pvwattsResult=o.pvwattsResult||null;globalSolarAtlasResult=o.globalSolarAtlasResult||null;gsaImages=o.gsaImages||[];renderGSAImages();renderPV();calc()}
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
async function captureGlobalSolarAtlas(){try{if(!parseCoordinates())throw Error('Vui lòng kiểm tra tọa độ.');if(!pid)throw Error('Trước tiên hãy lưu dự án bằng nút Lưu dự án.');let b=$('gsaCaptureBtn');b.disabled=true;b.textContent='Đang mở Atlas và chụp ảnh...';$('gsaStatus').textContent='Đang khởi chạy trình duyệt tự động. Có thể mất 1–2 phút.';let r=await api('/api/projects/'+pid+'/gsa-capture',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({lat:n('lat'),lon:n('lon')})});gsaImages=r.images||[];renderGSAImages();if(Number.isFinite(Number(r.ghi_kwh_m2_year))&&Number(r.ghi_kwh_m2_year)>0){$('solarRadiation').value=Number(r.ghi_kwh_m2_year).toFixed(1);$('gsaStatus').textContent='Đã chụp '+gsaImages.length+' ảnh. Đã tự điền GHI '+Number(r.ghi_kwh_m2_year).toFixed(1)+' kWh/m²/năm vào thông tin công trình.'}else{$('gsaStatus').textContent='Đã chụp '+gsaImages.length+' ảnh nhưng chưa đọc được giá trị GHI; bạn kiểm tra lại Map data.'}await listFiles()}catch(e){$('gsaStatus').textContent='Chưa hoàn tất: '+e.message}finally{$('gsaCaptureBtn').disabled=false;$('gsaCaptureBtn').textContent='Tự động lấy 4 hình ảnh'}}
function renderPV(){if(!pvwattsResult){$('pvResult').textContent='Chưa có kết quả mô phỏng.';$('pvChart').style.display='none';return}let a=pvwattsResult.monthly||[];let hourly=Array.isArray(pvwattsResult.hourly_average_kw)?pvwattsResult.hourly_average_kw:[];$('pvResult').innerHTML='<b>Sản lượng AC năm: '+Number(pvwattsResult.annual||0).toLocaleString('vi-VN',{maximumFractionDigits:0})+' kWh</b><br>'+a.map((v,i)=>['T1','T2','T3','T4','T5','T6','T7','T8','T9','T10','T11','T12'][i]+': '+Number(v).toLocaleString('vi-VN',{maximumFractionDigits:0})+' kWh').join(' · ')+'<br><b>PV trung bình 06:00–17:00: '+(hourly.length===24?hourly.slice(6,17).reduce((x,v)=>x+Number(v||0),0).toFixed(2)+' kWh/ngày':'Chưa có dữ liệu')+'</b><br><small>Trạm thời tiết: '+esc(pvwattsResult.station||'Chưa có thông tin')+' · '+esc(pvwattsResult.hourly_source||'')+'</small>';drawPVChart(a);if($('loadPvChart').style.display!=='none')drawLoadPvComparison()}
async function runPVWatts(){try{if(!parseCoordinates())throw Error('Vui lòng kiểm tra lại tọa độ công trình.');if(!azimuthConfirmed)throw Error('Bạn hãy nhấn Xác nhận hướng trên la bàn trước khi tính toán.');$('pvStatus').textContent='Đang lấy sản lượng tháng và 8.760 giá trị giờ từ PVWatts...';let body={lat:n('lat'),lon:n('lon'),system_capacity:n('pvSystemSize')||n('pv'),module_type:n('moduleType'),array_type:n('arrayType'),tilt:n('tilt'),azimuth:n('azimuth'),losses:n('losses'),dc_ac_ratio:n('dcAcRatio'),inv_eff:n('invEff'),gcr:n('gcr'),use_wf_albedo:$('albedoMode').value==='weather'?1:0};if($('albedoMode').value==='manual')body.albedo=n('albedo');let r=await api('/api/pvwatts',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});pvwattsResult=r;pvwattsResult._system_capacity_kw=Number(body.system_capacity)||0;$('annual').value=Math.round(r.annual||0);syncQuoteAuto();$('pvStatus').textContent='Mô phỏng thành công · PVWatts V'+r.version;renderPV();if($('loadSizingResults')?.style.display!=='none'&&typeof calculateLoadSizing==='function')calculateLoadSizing()}catch(e){$('pvStatus').textContent='Không lấy được dữ liệu: '+e.message}}

let loadSizingResult=null;

function calcBillNightShare(){let d=Math.max(0,Math.min(100,n('billDayShare')));$('billNightShare').value=(100-d).toFixed(1);if($('billNightShare').value.endsWith('.0'))$('billNightShare').value=(100-d).toFixed(0)}
function calculateResidentialBill(){
 try{
  const amount=Number($('residentialBillAmount').value);
  if(!Number.isFinite(amount)||amount<=0)throw Error('Nhập tổng tiền điện lớn hơn 0.');
  const vatIncluded=$('residentialVatIncluded').checked;
  const beforeVat=vatIncluded?amount/1.08:amount;
  const tiers=[
   {name:'Bậc 1',limit:100,rate:1901},
   {name:'Bậc 2',limit:100,rate:2289},
   {name:'Bậc 3',limit:200,rate:2894},
   {name:'Bậc 4',limit:300,rate:3456},
   {name:'Bậc 5',limit:Infinity,rate:3844}
  ];
  let remaining=beforeVat,kwh=0,html='<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse"><thead><tr><th style="text-align:left;padding:6px">Bậc</th><th style="text-align:right;padding:6px">Đơn giá</th><th style="text-align:right;padding:6px">Sản lượng</th><th style="text-align:right;padding:6px">Thành tiền</th></tr></thead><tbody>';
  tiers.forEach(t=>{
   const available=t.limit===Infinity?Infinity:t.limit;
   const costForTier=available===Infinity?remaining:available*t.rate;
   const spend=Math.min(remaining,costForTier);
   const units=spend/t.rate;
   html+='<tr><td style="padding:6px;border-top:1px solid #e2eae5">'+t.name+'</td><td style="padding:6px;text-align:right;border-top:1px solid #e2eae5">'+t.rate.toLocaleString('vi-VN')+'</td><td style="padding:6px;text-align:right;border-top:1px solid #e2eae5">'+units.toFixed(2)+'</td><td style="padding:6px;text-align:right;border-top:1px solid #e2eae5">'+spend.toLocaleString('vi-VN',{maximumFractionDigits:0})+'</td></tr>';
   kwh+=units;remaining-=spend;
  });
  const vat=beforeVat*.08,total=beforeVat+vat,daily=Number($('billDays').value)>0?kwh/Number($('billDays').value):0;
  html+='</tbody></table></div><div class="grid" style="margin-top:10px"><div class="gsa-card"><div class="muted">Tổng sản lượng ước tính</div><h2>'+kwh.toFixed(2)+'</h2><div class="muted">kWh</div></div><div class="gsa-card"><div class="muted">Tiền trước VAT</div><h2>'+beforeVat.toLocaleString('vi-VN',{maximumFractionDigits:0})+'</h2><div class="muted">đồng</div></div><div class="gsa-card"><div class="muted">VAT 8%</div><h2>'+vat.toLocaleString('vi-VN',{maximumFractionDigits:0})+'</h2><div class="muted">đồng</div></div><div class="gsa-card"><div class="muted">Tổng tiền thanh toán</div><h2 style="color:#16845b">'+Math.round(total).toLocaleString('vi-VN')+'</h2><div class="muted">đồng · đã gồm VAT</div></div></div>';
  $('residentialBillResult').innerHTML=html+'<p class="muted">Kết quả là ước tính theo biểu giá 5 bậc áp dụng từ 01/11/2026; giả định toàn bộ số tiền chỉ gồm tiền điện bậc thang và VAT 8%, chưa tính các khoản điều chỉnh khác nếu có.</p>';
  $('billMonthKwh').value=kwh.toFixed(2);
  $('billDailyAverage').value=Number($('billDays').value)>0?daily.toFixed(2):'';
  $('billAveragePrice').value=(amount/kwh).toFixed(2);syncQuoteAuto();
 }catch(e){$('residentialBillResult').textContent=e.message}
}
function updatePvModuleRecommendation(targetKwp){
 const box=$('resPvModules');if(!box)return;
 const wp=n('panelW')||625;
 if(!(targetKwp>0)||!(wp>0)){box.textContent='Nhập công suất tấm pin hợp lệ để quy đổi số lượng tấm.';return}
 const exact=targetKwp*1000/wp,low=Math.max(1,Math.floor(exact)),high=Math.max(1,Math.ceil(exact));
 const actualLow=low*wp/1000,actualHigh=high*wp/1000;
 const fmt=x=>Number(x).toLocaleString('vi-VN',{maximumFractionDigits:3});
 const option=(count,kwp,label)=>'<button type="button" class="btn alt" style="padding:5px 8px;font-size:11px;margin:3px 4px 0 0" onclick="choosePvModuleCount('+count+')">Chọn '+count+' tấm · '+fmt(kwp)+' kWp'+(label?' ('+label+')':'')+'</button>';
 box.innerHTML='Với tấm '+wp+' Wp: cần khoảng '+exact.toFixed(2)+' tấm. Công suất thực tế theo số tấm nguyên:<br>'+option(low,actualLow,'thấp hơn mục tiêu')+(high!==low?option(high,actualHigh,'đạt/vượt mục tiêu'):'')+'<div class="muted" style="margin-top:4px">Chọn số tấm phù hợp; công suất lắp đặt sẽ được cập nhật theo số tấm thực tế.</div>';
}
function choosePvModuleCount(count){
 $('panelQty').value=count;calc();
 const actual=n('pv');
 $('pvSystemSize').value=actual.toFixed(3);
 $('resPv').textContent=actual.toFixed(3);
 const pvLabel=$('resPv').parentElement.querySelector('.muted');if(pvLabel)pvLabel.textContent='kWp DC · công suất lắp đặt thực tế';
 $('loadSizingStatus').textContent='Đã chọn '+count+' tấm PV, tổng công suất lắp đặt '+actual.toFixed(3)+' kWp DC. Nếu muốn PVWatts mô phỏng theo công suất thực tế này, hãy bấm Tính toán PVWatts lại.';
}
function estimatePvKwpFromHourly(loadCoverage,nightSupply,household,fallbackDayKwh){
 const curve=pvwattsResult?.hourly_average_kw,simSize=Number(pvwattsResult?._system_capacity_kw)||0;
 if(!Array.isArray(curve)||curve.length!==24||simSize<=0)return null;
 const start=6,end=household?18:17,steps=(end-start)*2,loads=[];
 for(let i=0;i<steps;i++){
  const el=$('loadProfile'+i),raw=el?.value?.trim();
  let kw;
  if(raw!==undefined&&raw!=='')kw=Number(raw);
  else kw=household?(fallbackDayKwh/12):fallbackDayKwh/11;
  if(!Number.isFinite(kw)||kw<0)return null;
  loads.push(kw);
 }
 const solarLoad=loads.reduce((a,v)=>a+v*.5,0),requiredSelfUse=solarLoad*loadCoverage/100;
 const pvAt=h=>{const lo=Math.max(0,Math.min(23,Math.floor(h))),hi=Math.min(23,lo+1),f=h-lo;return Number(curve[lo]||0)*(1-f)+Number(curve[hi]||0)*f};
 const evaluate=kwp=>{
  let self=0,surplus=0,energy=0;
  for(let i=0;i<steps;i++){
   const h=start+(i+.5)/2,pv=pvAt(h)*kwp/simSize,load=loads[i];
   energy+=pv*.5;self+=Math.min(load,pv)*.5;surplus+=Math.max(0,pv-load)*.5;
  }
  return {self,surplus,energy};
 };
 let lo=0.1,hi=200;
 if(evaluate(hi).self+1e-6<requiredSelfUse||evaluate(hi).surplus+1e-6<nightSupply)return null;
 for(let k=0;k<45;k++){const mid=(lo+hi)/2,v=evaluate(mid);if(v.self>=requiredSelfUse&&v.surplus>=nightSupply)hi=mid;else lo=mid}
 const result=evaluate(hi);
 return {pvKwp:hi,solarWindowLoadKwh:solarLoad,requiredSelfUseKwh:requiredSelfUse,pvSelfUseKwh:result.self,pvSurplusKwh:result.surplus,pvEnergyKwh:result.energy,simulationSizeKwp:simSize};
}
function calculateLoadSizing(){
 try{
  const monthly=n('billMonthKwh'),days=n('billDays'),dayShare=n('billDayShare'),nightShare=100-dayShare;
  if(monthly<=0||days<=0)throw new Error('Nhập tổng điện năng hóa đơn và số ngày hợp lệ.');
  if(dayShare<0||dayShare>100)throw new Error('Tỷ lệ ban ngày phải từ 0 đến 100%.');
  const psh=n('billPsh'),pr=0.80,pvCoverage=0.80,nightCoverage=n('billNightCoverage')/100,dod=n('billDod')/100,chargeEff=n('billChargeEff')/100,dischargeEff=n('billDischargeEff')/100;
  if(nightCoverage<0||nightCoverage>1)throw new Error('Tỷ lệ pin đáp ứng điện ban đêm phải từ 0 đến 100%.');
  const daily=monthly/days,day=daily*dayShare/100,night=daily-day;
  if(psh<=0||pr<=0||dod<=0||chargeEff<=0||chargeEff>1||dischargeEff<=0||dischargeEff>1)throw new Error('Kiểm tra PSH, DoD và hiệu suất sạc/xả pin (50–100%).');
  const nightSupply=night*nightCoverage;
  // NightSupply is energy delivered by the battery to loads. Account for discharge
  // losses to determine stored energy, then charge losses to determine PV surplus.
  const batteryStoredRequired=nightSupply/dischargeEff;
  const pvEnergyForBattery=batteryStoredRequired/chargeEff;
  const hourlySizing=estimatePvKwpFromHourly(pvCoverage*100,pvEnergyForBattery,$('loadCustomerMode')?.value==='household',day);
  const pvEnergyForDay=hourlySizing?hourlySizing.requiredSelfUseKwh:day*pvCoverage;
  const totalPvEnergy=pvEnergyForDay+pvEnergyForBattery;
  const pv=hourlySizing?hourlySizing.pvKwp:totalPvEnergy/(psh*pr);
  const battery=nightSupply/(dod*dischargeEff);
  loadSizingResult={monthlyKwh:monthly,billDays:days,dayShare,nightShare,dailyKwh:daily,dayKwh:day,nightKwh:night,pvCoverage:pvCoverage*100,nightCoverage:nightCoverage*100,psh,pr:pr*100,dod:dod*100,chargeEfficiency:chargeEff*100,dischargeEfficiency:dischargeEff*100,batteryStoredRequiredKwh:batteryStoredRequired,pvEnergyForDayKwh:pvEnergyForDay,pvEnergyForBatteryKwh:pvEnergyForBattery,totalPvEnergyKwh:totalPvEnergy,pvKwp:pv,nightSupplyKwh:nightSupply,batteryKwh:battery,pvSizingMethod:hourlySizing?'PVWatts hourly average + load profile':'PSH / PR estimate',hourlySizing};
  $('resDaily').textContent=daily.toFixed(1);$('resDay').textContent=day.toFixed(1);$('resNight').textContent=night.toFixed(1);$('resPv').textContent=pv.toFixed(1);updatePvModuleRecommendation(pv);$('resBattery').textContent=battery.toFixed(1);$('resNightSupply').textContent=nightSupply.toFixed(1);
  $('loadSizingAssumptions').textContent=(hourlySizing?'Sizing theo đường PVWatts giờ: cần tự cấp tối thiểu '+pvEnergyForDay.toFixed(1)+' kWh trong khung có nắng và còn dư '+pvEnergyForBattery.toFixed(1)+' kWh để nạp pin. Đã dùng phụ tải '+(hourlySizing.solarWindowLoadKwh).toFixed(1)+' kWh trong khung '+( $('loadCustomerMode')?.value==='household'?'06:00–18:00':'06:00–17:00')+' và chuẩn hóa đường PVWatts từ hệ '+hourlySizing.simulationSizeKwp.toFixed(2)+' kWp.':'Tính sơ bộ theo PSH/PR: '+pvEnergyForDay.toFixed(1)+' kWh/ngày cấp tải ban ngày và '+pvEnergyForBattery.toFixed(1)+' kWh/ngày để sạc pin; tổng nhu cầu PV '+totalPvEnergy.toFixed(1)+' kWh/ngày. ')+'Giả định: PSH '+psh.toFixed(1)+' giờ/ngày; PR '+(pr*100).toFixed(0)+'%; đáp ứng ban ngày '+(pvCoverage*100).toFixed(0)+'%; đáp ứng ban đêm '+(nightCoverage*100).toFixed(0)+'%; hiệu suất sạc pin '+(chargeEff*100).toFixed(0)+'%; hiệu suất xả pin '+(dischargeEff*100).toFixed(0)+'%; DoD '+(dod*100).toFixed(0)+'%. Kết quả theo PVWatts dùng đường sản lượng trung bình năm, cần kiểm tra tháng sản lượng thấp và giới hạn inverter trước khi chốt thiết kế.';
  $('loadSizingResults').style.display='block';$('loadSizingStatus').textContent=hourlySizing?'Đã tự tối ưu công suất bằng PVWatts và phụ tải 30 phút: PV dư cần '+pvEnergyForBattery.toFixed(2)+' kWh/ngày để pin cấp '+nightSupply.toFixed(2)+' kWh/ngày cho tải.':'Chưa có đường PVWatts hợp lệ hoặc chưa dựng phụ tải 30 phút; đang dùng ước tính PSH/PR. Chạy PVWatts và dựng phụ tải rồi tính lại để tối ưu tự động.';return loadSizingResult;
 }catch(e){$('loadSizingStatus').textContent=e.message;loadSizingResult=null;return null}
}
let targetBatteryPvResult=null;
function calculateTargetBatteryPv(){
 try{
  if(!loadSizingResult&&!calculateLoadSizing())return null;
  const target=Number($('targetBatteryDailyKwh').value),dayCoverage=Number($('targetPvDayCoverage').value),base=loadSizingResult;
  if(!Number.isFinite(target)||target<0)throw Error('Nhập điện năng muốn nạp vào pin mỗi ngày hợp lệ.');
  if(!Number.isFinite(dayCoverage)||dayCoverage<0||dayCoverage>100)throw Error('Tỷ lệ PV cấp tải ban ngày phải từ 0 đến 100%.');
  const pvDay=base.dayKwh*dayCoverage/100;
  const pvCharge=target; // Charging efficiency is fixed at 100%.
  const totalPvEnergy=pvDay+pvCharge;
  const pvKwp=totalPvEnergy/(base.psh*(base.pr/100));
  const dod=base.dod/100;
  if(dod<=0)throw Error('DoD phải lớn hơn 0%.');
  const batteryNominal=target/dod;
  targetBatteryPvResult={...base,pvCoverage:dayCoverage,pvEnergyForDayKwh:pvDay,pvEnergyForBatteryKwh:pvCharge,totalPvEnergyKwh:totalPvEnergy,pvKwp,batteryKwh:batteryNominal,nightSupplyKwh:target};updatePvModuleRecommendation(pvKwp);
  $('targetBatteryPvResult').innerHTML='<strong>PV đề xuất: '+pvKwp.toFixed(1)+' kWp DC</strong><br>Dung lượng pin danh định: '+batteryNominal.toFixed(1)+' kWh<br>PV cấp tải ban ngày: '+pvDay.toFixed(1)+' kWh/ngày · Nạp pin: '+pvCharge.toFixed(1)+' kWh/ngày · Tổng PV: '+totalPvEnergy.toFixed(1)+' kWh/ngày<br><small>Giả định PSH '+base.psh.toFixed(1)+' giờ/ngày, PR '+base.pr.toFixed(0)+'%, hiệu suất sạc 100%, DoD '+base.dod.toFixed(0)+'%.</small>';
  return targetBatteryPvResult;
 }catch(e){$('targetBatteryPvResult').textContent=e.message;targetBatteryPvResult=null;return null}
}
function applyTargetBatteryPv(){
 const r=calculateTargetBatteryPv();if(!r)return;
 loadSizingResult=r;applyLoadSizing();
 $('loadSizingStatus').textContent='Đã áp dụng phương án PV theo dung lượng pin mục tiêu. Kiểm tra cấu hình inverter và lưu dự án.';
}
let loadProfileTimes=Array.from({length:23},(_,i)=>{const mins=360+i*30;return String(Math.floor(mins/60)).padStart(2,'0')+':'+String(mins%60).padStart(2,'0')});
function setLoadProfileRange(){const household=$('loadCustomerMode')?.value==='household',next=Array.from({length:household?29:23},(_,i)=>{const mins=360+i*30;return String(Math.floor(mins/60)).padStart(2,'0')+':'+String(mins%60).padStart(2,'0')});const old={};loadProfileTimes.forEach((t,i)=>{const el=$('loadProfile'+i);if(el)old[t]=el.value});loadProfileTimes=next;window.__solarInitStage='before-initLoadProfileInputs';initLoadProfileInputs();window.__solarInitStage='after-initLoadProfileInputs';loadProfileTimes.forEach((t,i)=>{if(old[t]!==undefined)$('loadProfile'+i).value=old[t]});}
function updateTouScheduleDescription(){
 const mode=$('touScheduleMode')?.value||'old';
 $('touScheduleDescription').textContent=mode==='old'
 ?'Khung giờ cũ: Cao điểm Thứ Hai–Thứ Bảy 17:30–22:30; bình thường 06:00–17:30 và 22:30–24:00 (Chủ nhật 06:00–24:00); thấp điểm tất cả các ngày 00:00–06:00.'
 :'Khung giờ mới: Cao điểm Thứ Hai–Thứ Bảy 09:30–11:30 và 17:00–20:00; bình thường 04:00–09:30, 11:30–17:00 và 20:00–22:00 (Chủ nhật 04:00–22:00); thấp điểm tất cả các ngày 22:00–04:00 hôm sau.';
}
function toMinutes(v){if(!v||!v.includes(':'))return NaN;const p=v.split(':').map(Number);return p[0]*60+p[1]}
function toggleLoadCustomerMode(){
 const household=$('loadCustomerMode')?.value==='household';
 $('householdLoadInfo').style.display=household?'block':'none';
 $('factoryLoadFields').style.display=household?'none':'block';
 setLoadProfileRange();
}
function togglePlantScheduleMode(){
 const continuous=$('plantLoadMode').value==='continuous';
 $('plantScheduleFields').style.display=continuous?'none':'block';
}
function buildLoadProfileFromTou(){
 try{
  const group=$('billGroupSelect').value;
  const rows=group?billBatchResults.filter(x=>x._included&&billGroupKey(x)===group):[];
  const household=$('loadCustomerMode')?.value==='household';
  if(household){
   // Household mode can use either the selected invoice batch or the
   // already-populated total kWh / billing days fields. Do not require a
   // commercial tariff group for residential bills or manually entered data.
   let totalKwh=0,totalDays=0,sourceLabel='';
   if(rows.length){
    rows.forEach(row=>{const st=new Date(row.period_start+'T00:00:00Z'),en=new Date(row.period_end+'T00:00:00Z');if(isNaN(st)||isNaN(en)||en<st)throw Error('Ngày kỳ hóa đơn không hợp lệ: '+(row.filename||''));totalDays+=Math.round((en-st)/86400000)+1;totalKwh+=Number(row.monthly_kwh)||0});
    sourceLabel=rows.length+' hóa đơn / '+totalDays+' ngày';
   }else{
    totalKwh=Number($('billMonthKwh').value)||0;
    totalDays=Number($('billDays').value)||0;
    if(totalKwh<=0||totalDays<=0)throw Error('Chưa có tổng kWh và số ngày hợp lệ. Hãy nhập tổng điện năng hóa đơn và số ngày, hoặc chọn nhóm hóa đơn đã đọc.');
    sourceLabel='tổng hóa đơn đã nhập / '+totalDays+' ngày';
   }
   const dayShare=Number($('billDayShare').value),nightShare=Number($('billNightShare').value);
   if(!Number.isFinite(dayShare)||dayShare<0||dayShare>100||!Number.isFinite(nightShare)||nightShare<0||nightShare>100||Math.abs(dayShare+nightShare-100)>.11)throw Error('Tỷ lệ điện năng ban ngày 06:00–18:00 và ban đêm 18:00–06:00 phải hợp lệ và tổng bằng 100%.');
   if(totalDays<=0||totalKwh<=0)throw Error('Không có tổng kWh hoặc số ngày hợp lệ để dựng phụ tải hộ dân.');
   const dailyKwh=totalKwh/totalDays,dayKwh=dailyKwh*dayShare/100,nightKwh=dailyKwh*nightShare/100,dayKw=dayKwh/12,nightKw=nightKwh/12;
   loadProfileTimes.forEach((t,i)=>{$('loadProfile'+i).value=(Number(t.slice(0,2))<18?dayKw:nightKw).toFixed(2)});
   $('loadSizingStatus').textContent='Đã dựng phụ tải hộ dân từ '+sourceLabel+'. Tổng trung bình '+dailyKwh.toFixed(2)+' kWh/ngày; ban ngày 06:00–18:00: '+dayShare.toFixed(1)+'% ('+dayKwh.toFixed(2)+' kWh/ngày, '+dayKw.toFixed(2)+' kW); ban đêm 18:00–06:00: '+nightShare.toFixed(1)+'% ('+nightKwh.toFixed(2)+' kWh/ngày, '+nightKw.toFixed(2)+' kW). Biểu đồ vẫn hiển thị 06:00–20:00; 00:00–06:00 được tính vào nhu cầu pin.';
   drawLoadPvComparison();return;
  }
  if(!group)throw Error('Hãy chọn nhóm hóa đơn trước.');
  if(!rows.length)throw Error('Nhóm hóa đơn đã chọn không có dữ liệu.');
  if(!rows.every(x=>['kinh doanh','sản xuất'].some(t=>(x.customer_type||'').toLowerCase().includes(t))))throw Error('Nhà máy cần chọn nhóm hóa đơn điện sản xuất/kinh doanh.');
  if(rows.some(x=>!x.period_start||!x.period_end))throw Error('Hóa đơn thiếu ngày bắt đầu/kết thúc kỳ.');
  const continuous=$('plantLoadMode').value==='continuous';
  if(!continuous&&rows.some(x=>!x.tou_kwh||!['normal','peak','offpeak'].every(k=>Number.isFinite(Number(x.tou_kwh[k])))))throw Error('Chế độ theo lịch ca cần đủ kWh bình thường, cao điểm và thấp điểm.');
  if(continuous){
   let totalKwh=0,totalDays=0;
   rows.forEach(row=>{const st=new Date(row.period_start+'T00:00:00Z'),en=new Date(row.period_end+'T00:00:00Z');if(!row.period_start||!row.period_end||isNaN(st)||isNaN(en)||en<st)throw Error('Hóa đơn thiếu hoặc sai ngày kỳ: '+(row.filename||''));totalDays+=Math.round((en-st)/86400000)+1;totalKwh+=Number(row.monthly_kwh)||0});
   if(totalDays<=0||totalKwh<=0)throw Error('Không có tổng kWh hoặc số ngày hợp lệ để dựng phụ tải 24/7.');
   const flatKw=totalKwh/totalDays/24;
   loadProfileTimes.forEach((t,i)=>{$('loadProfile'+i).value=flatKw.toFixed(2)});
   const energy0617=flatKw*11;
   $('loadSizingStatus').textContent='Đã dựng phụ tải liên tục 24/7 từ '+rows.length+' hóa đơn / '+totalDays+' ngày. Phụ tải trung bình cố định '+flatKw.toFixed(2)+' kW cả ngày và đêm; điện 06:00–17:00 ước tính '+energy0617.toFixed(1)+' kWh/ngày. Không áp dụng lịch ca, nghỉ trưa hoặc tải nền ngoài ca. Đây là mô hình giả định phụ tải phẳng, cần xác nhận bằng dữ liệu đo thực tế.';
   drawLoadPvComparison();return;
  }
  const startMin=toMinutes($('plantStart').value),endMin=toMinutes($('plantEnd').value),lunchEnabled=$('plantLunchEnabled').checked,lunchStart=toMinutes($('plantLunchStart').value),lunchEnd=toMinutes($('plantLunchEnd').value);
  const workDays=Number($('plantWorkDays').value),lunchRaw=$('plantLunchOff').value.trim(),outsideRaw=$('plantOutsideShare').value.trim(),lunchOff=lunchEnabled?Number(lunchRaw)/100:0,outsideFactor=Number(outsideRaw)/100;
  if(!(startMin>=0&&endMin>startMin&&endMin<=1440))throw Error('Giờ bắt đầu/kết thúc ca không hợp lệ.');
  if(lunchEnabled&&!(lunchStart>=startMin&&lunchEnd>lunchStart&&lunchEnd<=endMin))throw Error('Giờ nghỉ trưa phải nằm trong ca làm việc.');
  if((lunchEnabled&&lunchRaw==='')||outsideRaw===''||!Number.isFinite(lunchOff)||lunchOff<0||lunchOff>1||!Number.isFinite(outsideFactor)||outsideFactor<0||outsideFactor>1)throw Error('Vui lòng kiểm tra tỷ lệ máy dừng giờ nghỉ và tải nền ngoài ca (0–100%).');
  let totalDays=0,normalKwh=0,peakKwh=0,offKwh=0;const dates=[];
  const parseDate=v=>{const m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);return m?new Date(Date.UTC(+m[1],+m[2]-1,+m[3])):null};
  rows.forEach(row=>{const st=parseDate(row.period_start),en=parseDate(row.period_end);if(!st||!en||en<st)throw Error('Ngày kỳ hóa đơn không hợp lệ: '+(row.filename||''));for(let d=new Date(st);d<=en;d.setUTCDate(d.getUTCDate()+1)){dates.push(new Date(d));totalDays++}normalKwh+=Number(row.tou_kwh.normal)||0;peakKwh+=Number(row.tou_kwh.peak)||0;offKwh+=Number(row.tou_kwh.offpeak)||0});
  const mode=$('touScheduleMode').value;
  const isWorking=date=>{const dow=date.getUTCDay();return workDays===7||(workDays===6&&dow!==0)||(workDays===5&&dow>=1&&dow<=5)};
  const tariff=(date,minute)=>{
   const sunday=date.getUTCDay()===0;
   if(mode==='new'){if(!sunday&&((minute>=570&&minute<690)||(minute>=1020&&minute<1200)))return 'peak';if(minute>=1320||minute<240)return 'offpeak';return 'normal'}
   if(!sunday&&minute>=1050&&minute<1350)return 'peak';if(minute<360)return 'offpeak';return 'normal';
  };
  // Convert bill kWh into average kW for each tariff class using all actual calendar days in the billing periods.
  const eff={normal:0,peak:0,offpeak:0};
  dates.forEach(date=>{const working=isWorking(date);for(let m=0;m<1440;m+=30){const cls=tariff(date,m);let factor=outsideFactor;if(working&&m>=startMin&&m<endMin){const overlap=lunchEnabled?Math.max(0,Math.min(m+30,lunchEnd)-Math.max(m,lunchStart))/30:0;factor=1-overlap*lunchOff}eff[cls]+=.5*factor}});
  const energy={normal:normalKwh,peak:peakKwh,offpeak:offKwh},kw={};
  for(const k of ['normal','peak','offpeak']){if(energy[k]>0&&eff[k]<=0)throw Error('Không đủ giờ quy đổi cho khung '+k+'. Kiểm tra lịch vận hành và tải nền ngoài ca.');kw[k]=eff[k]>0?energy[k]/eff[k]:0}
  // Build a representative week (Mon-Sun), then average all seven days.
  // This correctly treats Saturday as off for a five-day factory and Sunday as off for a six-day factory.
  const week=Array.from({length:7},(_,i)=>new Date(Date.UTC(2026,0,5+i)));
  loadProfileTimes.forEach((t,i)=>{
   const minute=360+i*30;
   const daily=week.map(date=>{
    const working=isWorking(date),cls=tariff(date,minute);
    if(!working||minute<startMin||minute>=endMin)return kw[cls]*outsideFactor;
    const overlap=lunchEnabled?Math.max(0,Math.min(minute+30,lunchEnd)-Math.max(minute,lunchStart))/30:0;
    return kw[cls]*(1-overlap*lunchOff);
   });
   $('loadProfile'+i).value=(daily.reduce((sum,v)=>sum+v,0)/7).toFixed(2);
  });
  const energy0617=loadProfileTimes.slice(0,-1).reduce((sum,_,i)=>sum+(Number($('loadProfile'+i).value)+Number($('loadProfile'+(i+1)).value))*.25,0);
  $('loadSizingStatus').textContent='Đã tự dựng phụ tải 30 phút từ '+rows.length+' hóa đơn / '+totalDays+' ngày. Lịch '+workDays+' ngày/tuần; ca '+$('plantStart').value+'–'+$('plantEnd').value+'; '+(lunchEnabled?'nghỉ trưa '+$('plantLunchStart').value+'–'+$('plantLunchEnd').value+', dừng '+$('plantLunchOff').value+'% máy':'không áp dụng nghỉ trưa')+'; tải nền ngoài ca/ngày nghỉ '+$('plantOutsideShare').value+'%. Điện 06:00–17:00 ước tính theo ngày trung bình tuần: '+energy0617.toFixed(1)+' kWh. Đây là mô hình ước tính từ hóa đơn, không phải số đo công tơ.';
  drawLoadPvComparison();
 }catch(e){$('loadSizingStatus').textContent=e.message}
}
function toggleLoadSizing(){
 const content=$('loadSizingContent'),button=$('loadSizingToggle');
 const opening=content.style.display==='none';
 content.style.display=opening?'block':'none';
 button.textContent=opening?'− Thu gọn phần tính toán':'＋ Mở phần tính toán';
 button.setAttribute('aria-expanded',opening?'true':'false');
}
function initLoadProfileInputs(){
 const box=$('loadProfileInputs');if(!box)return;
 box.innerHTML=loadProfileTimes.map((t,i)=>'<label style="font-size:11px;font-weight:600">'+t+'<input id="loadProfile'+i+'" type="number" min="0" step="0.1" value="" placeholder="kW" style="margin-top:4px;padding:8px"></label>').join('');
}
function clearLoadProfile(){loadProfileTimes.forEach((_,i)=>$('loadProfile'+i).value='');$('loadPvChart').style.display='none'}
function fillFlatLoadProfile(){
 const r=loadSizingResult||calculateLoadSizing();if(!r)return;
 const flat=r.dayKwh/($('loadCustomerMode')?.value==='household'?14:11);
 loadProfileTimes.forEach((_,i)=>$('loadProfile'+i).value=flat.toFixed(1));
 drawLoadPvComparison();
}
function drawLoadPvComparison(){
 try{
  const household=$('loadCustomerMode')?.value==='household',startHour=6,endHour=household?20:17,intervals=(endHour-startHour)*2;
  const load=loadProfileTimes.map((t,i)=>{const raw=$('loadProfile'+i).value.trim();if(raw==='')throw Error('Vui lòng nhập đủ phụ tải tại các mốc 30 phút.');const v=Number(raw);if(!Number.isFinite(v)||v<0)throw Error('Phụ tải phải là số không âm tại mốc '+t+'.');return v});
  const energy=load.slice(0,intervals).reduce((sum,v)=>sum+v*.5,0),peak=Math.max(...load),average=energy/(endHour-startHour);
  const hourly=pvwattsResult?.hourly_average_kw,hasPv=Array.isArray(hourly)&&hourly.length===24;
  // Interpolate the PVWatts hourly average curve at each 30-minute interval midpoint.
  // This makes battery discharge begin at the actual PV/load crossover, not at the next whole hour.
  const pvAt=i=>{if(!hasPv)return 0;const h=startHour+(i+.5)/2,lo=Math.floor(h),hi=Math.min(23,lo+1),f=h-lo;return Number(hourly[lo]||0)*(1-f)+Number(hourly[hi]||0)*f};
  const pvEnergy=hasPv?hourly.slice(startHour,endHour).reduce((sum,v)=>sum+Number(v||0),0):0;
  let selfUse=0,gridImport=0,exportEnergy=0,batteryDischarge=0,batteryExhaustAt=null,daytimeLoad6to17=0,daytimePv6to17=0,daytimeSelfUse6to17=0;
  const batteryCurve=Array(intervals).fill(0),batteryFillRatio=Array(intervals).fill(0);
  if(hasPv){
   let remaining=0;
   if(household){const total=Number($('batTotal')?.value)||((Number($('batKwh')?.value)||loadSizingResult?.batteryKwh||0)*(Number($('batQty')?.value)||1));const dod=(Number($('billDod')?.value)||90)/100;remaining=Math.max(0,total*dod)}
   for(let i=0;i<intervals;i++){
    const le=load[i]*.5,pe=pvAt(i)*.5,used=Math.min(le,pe),surplus=Math.max(0,pe-le),deficit=Math.max(0,le-pe);
    selfUse+=used;exportEnergy+=surplus;
    if(i<22){daytimeLoad6to17+=le;daytimePv6to17+=pe;daytimeSelfUse6to17+=used;}
    let discharge=0;
    if(household&&deficit>0&&remaining>0){discharge=Math.min(deficit,remaining);remaining-=discharge;batteryDischarge+=discharge;if(remaining<=1e-6&&batteryExhaustAt===null)batteryExhaustAt=startHour+(i+1)/2}
    batteryCurve[i]=discharge/.5;
    batteryFillRatio[i]=deficit>0?Math.min(1,discharge/deficit):0;
    gridImport+=Math.max(0,deficit-discharge);
   }
  }
  const W=900,H=340,L=58,R=18,T=24,B=42,pw=W-L-R,ph=H-T-B;
  const batteryKw=batteryCurve.map(v=>v),max=Math.max(1,...load,...(hasPv?hourly.slice(startHour,endHour+1).map(Number):[]),...batteryKw)*1.12;
  const x=i=>L+pw*i/intervals,y=v=>T+ph*(1-v/max);
  let grid='',labels='';
  for(let k=0;k<=4;k++){const val=max*k/4,yy=y(val);grid+='<line x1="'+L+'" y1="'+yy+'" x2="'+(W-R)+'" y2="'+yy+'" stroke="#e4ece8"/><text x="'+(L-9)+'" y="'+(yy+4)+'" text-anchor="end" fill="#687a73" font-size="11">'+val.toFixed(1)+'</text>'}
  loadProfileTimes.forEach((t,i)=>{if(i%2===0||i===intervals)labels+='<text x="'+x(i)+'" y="'+(H-13)+'" text-anchor="middle" fill="#52665d" font-size="11">'+t+'</text>'});
  let step='M'+x(0).toFixed(1)+' '+y(load[0]).toFixed(1);
  for(let i=1;i<load.length;i++)step+=' L'+x(i).toFixed(1)+' '+y(load[i-1]).toFixed(1)+' L'+x(i).toFixed(1)+' '+y(load[i]).toFixed(1);
  let pvSmooth='',pts=[];
  if(hasPv){pts=hourly.slice(startHour,endHour+1).map((v,i)=>({x:x(i*2),y:y(Number(v||0))}));pvSmooth='M'+pts[0].x.toFixed(1)+' '+pts[0].y.toFixed(1);for(let i=0;i<pts.length-1;i++){const p0=pts[Math.max(0,i-1)],p1=pts[i],p2=pts[i+1],p3=pts[Math.min(pts.length-1,i+2)],c1x=p1.x+(p2.x-p0.x)/6,c1y=p1.y+(p2.y-p0.y)/6,c2x=p2.x-(p3.x-p1.x)/6,c2y=p2.y-(p3.y-p1.y)/6;pvSmooth+=' C'+c1x.toFixed(1)+' '+c1y.toFixed(1)+','+c2x.toFixed(1)+' '+c2y.toFixed(1)+','+p2.x.toFixed(1)+' '+p2.y.toFixed(1)}}
  const pvArea=hasPv&&pts.length===endHour-startHour+1?pvSmooth+' L'+pts[pts.length-1].x.toFixed(1)+' '+y(0).toFixed(1)+' L'+pts[0].x.toFixed(1)+' '+y(0).toFixed(1)+' Z':'';
  // Draw only the actual missing-power gap: PV is the lower boundary and
  // the battery-supported load is the upper boundary, capped at the load.
  // Match the exact Catmull-Rom curve used by the green PV line, so the blue
  // battery area touches PV continuously without thin white wedges.
  const pvBoundary=i=>{
   if(!hasPv)return 0;
   const h=startHour+i/2,base=Math.floor(h),t=h-base;
   const val=k=>Number(hourly[Math.max(startHour,Math.min(endHour,k))]||0);
   const p0=val(base-1),p1=val(base),p2=val(base+1),p3=val(base+2);
   return Math.max(0,.5*((2*p1)+(-p0+p2)*t+(2*p0-5*p1+4*p2-p3)*t*t+(-p0+3*p1-3*p2+p3)*t*t*t));
  };
  let batteryPath='';
  for(let i=0;i<intervals;i++){
   const p0=pvBoundary(i),p1=pvBoundary(i+1),ratio=batteryFillRatio[i];
   const gap0=Math.max(0,load[i]-p0),gap1=Math.max(0,load[i]-p1);
   const top0=p0+gap0*ratio,top1=p1+gap1*ratio;
   if(ratio>0.001&&Math.max(top0-p0,top1-p1)>0.001){
    batteryPath+=' M'+x(i).toFixed(2)+' '+y(top0).toFixed(2)+' L'+x(i+1).toFixed(2)+' '+y(top1).toFixed(2)+' L'+x(i+1).toFixed(2)+' '+y(p1).toFixed(2)+' L'+x(i).toFixed(2)+' '+y(p0).toFixed(2)+' Z';
   }
  }
  $('loadPvChartSvg').innerHTML='<svg viewBox="0 0 '+W+' '+H+'" style="width:100%;height:auto;font-family:Arial,sans-serif" role="img" aria-label="Biểu đồ PV, phụ tải và công suất xả pin từ 6 giờ đến '+endHour+' giờ">'+grid+(hasPv?'<path d="'+pvArea+'" fill="#35a96f" fill-opacity=".28" stroke="none"/><path d="'+pvSmooth+'" fill="none" stroke="#168a55" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>':'')+'<path d="'+step+'" fill="none" stroke="#e53935" stroke-width="3.5" stroke-linejoin="miter" stroke-linecap="square"/>'+(household&&hasPv?'<path d="'+batteryPath+'" fill="#2878d0" fill-opacity=".30" stroke="none"/>':'')+labels+'<text x="15" y="16" fill="#687a73" font-size="11">kW</text></svg>';
  const card=(label,value,unit)=>'<div class="gsa-card"><div class="muted">'+label+'</div>'+(unit==='%'?'<h3 style="margin:7px 0;color:#16845b;white-space:nowrap">'+value.toFixed(1)+'<span style="font-size:.78em;margin-left:3px">%</span></h3>':'<h3 style="margin:7px 0;color:#16845b">'+value.toFixed(1)+'</h3><div class="muted">'+unit+'</div>')+'</div>';
  let metrics=card('Điện năng phụ tải '+startHour+'–'+endHour,energy,'kWh/ngày')+card('Công suất phụ tải lớn nhất',peak,'kW')+card('Công suất phụ tải trung bình '+startHour+'–'+endHour,average,'kW');
  if(hasPv){const loadCoverage=daytimeLoad6to17>0?daytimeSelfUse6to17/daytimeLoad6to17*100:0,pvSelfRate=daytimePv6to17>0?daytimeSelfUse6to17/daytimePv6to17*100:0;metrics+=card('PV trung bình '+startHour+'–'+endHour,pvEnergy,'kWh/ngày')+card('PV tự tiêu thụ',selfUse,'kWh/ngày')+card('Điện lưới cần bổ sung',gridImport,'kWh/ngày')+card('PV dư / có thể sạc vào pin lưu trữ',exportEnergy,'kWh/ngày')+card('Tỷ lệ phụ tải được PV đáp ứng (06–17)',loadCoverage,'%')+card('Tỷ lệ PV được tự tiêu thụ (06–17)',pvSelfRate,'%');if(household)metrics+=(batteryExhaustAt!==null?'<div class="gsa-card"><div class="muted">Pin cạn theo mô hình</div><h3 style="margin:7px 0;color:#2878d0">'+batteryExhaustAt.toFixed(1).replace('.',':')+' h</h3><div class="muted">Sau thời điểm này lưới cấp phần thiếu</div></div>':'<div class="gsa-card"><div class="muted">Trạng thái pin lúc 20:00</div><h3 style="margin:7px 0;color:#2878d0">Còn năng lượng</h3><div class="muted">Pin có thể tiếp tục xả sau 20:00</div></div>')}
  $('loadPvMetrics').innerHTML=metrics;
  if($('batteryLegend'))$('batteryLegend').style.display=household&&hasPv?'inline':'none';
  $('loadPvNote').textContent=hasPv?(household?'Miền xanh: PV; đường đỏ: phụ tải hộ dân; đường xanh dương: công suất xả pin. Pin chỉ bắt đầu xả khi PV thấp hơn tải. Mô hình giới hạn năng lượng xả theo tổng dung lượng pin × DoD; nếu pin cạn trước 20:00, đường xả về 0 và điện lưới bù phần thiếu. Nếu còn pin lúc 20:00, biểu đồ dừng tại 20:00 nhưng pin được hiểu là tiếp tục xả sau thời điểm này.':'Miền xanh: PV; đường đỏ: phụ tải nhà máy. Dữ liệu PV lấy trung bình 8.760 giờ của năm khí tượng điển hình.'):'Chưa có dữ liệu PVWatts hourly. Hãy chạy PVWatts để hiển thị đường sản lượng PV.';
  $('loadPvChart').style.display='block';
 }catch(e){$('loadSizingStatus').textContent=e.message}
}
function applyLoadSizing(){
 try{
  const r=loadSizingResult||calculateLoadSizing();
  if(!r)return;
  $('load').value=Number(r.monthlyKwh.toFixed(2));
  const panelW=n('panelW')||625;$('panelQty').value=Math.max(1,Math.ceil(r.pvKwp*1000/panelW));
  $('batKwh').value=r.batteryKwh.toFixed(1);$('batQty').value=1;
  $('usagePattern').value='Tỷ lệ tiêu thụ ban ngày '+r.dayShare.toFixed(1)+'%, ban đêm '+r.nightShare.toFixed(1)+'%; điện năng trung bình ban ngày '+r.dayKwh.toFixed(1)+' kWh/ngày, ban đêm '+r.nightKwh.toFixed(1)+' kWh/ngày.';
  calc();$('loadSizingStatus').textContent='Đã áp dụng điện năng tháng, số lượng tấm PV dự kiến và dung lượng pin vào cấu hình. Hãy kiểm tra và lưu dự án.';
 }catch(e){$('loadSizingStatus').textContent=e.message}
}
let billBatchResults=[];
function billFmt(v,options={}){return v===null||v===undefined||v===''?'—':Number(v).toLocaleString('vi-VN',{maximumFractionDigits:2,...options})}
function billPeriodKey(x){return x.period_start&&x.period_end?x.period_start+'|'+x.period_end:null}
function billGroupKey(x){return [x.customer_type||'Chưa xác định',x.voltage||'Chưa xác định cấp áp',x.tariff_type||'Chưa xác định hình thức giá'].join(' | ')}
function applyBillGroup(){
 const key=$('billGroupSelect').value;const rows=billBatchResults.filter(x=>x._included&&billGroupKey(x)===key);
 if(!rows.length)return;
 const kwh=rows.reduce((a,x)=>a+(Number(x.monthly_kwh)||0),0),days=rows.reduce((a,x)=>a+(Number(x.period_days)||0),0);
 $('billMonthKwh').value=kwh.toFixed(2);$('billDays').value=days||30;$('billDailyAverage').value=days>0?(kwh/days).toFixed(2):'';const money=rows.reduce((a,x)=>a+(Number(x.amount_vnd)||0),0);$('billAveragePrice').value=kwh>0&&money>0?(money/kwh).toFixed(2):'';const normalizedMonthFactor=days>0?30/days:0;$('billMonthlyAverageKwh').value=days>0?billFmt(kwh*normalizedMonthFactor):'';$('billMonthlyAverageMoney').value=days>0?billFmt(money*normalizedMonthFactor,{maximumFractionDigits:0}):'';syncQuoteAuto();loadSizingResult=null;
 $('billReadStatus').textContent='Đã chọn nhóm '+key+' để tính sizing. Các nhóm khác không bị cộng chung.';
 if(rows.every(x=>x.period_start&&x.period_end&&x.tou_kwh&&['normal','peak','offpeak'].every(k=>Number.isFinite(Number(x.tou_kwh[k])))))buildLoadProfileFromTou();
}
function updateBillManual(index,key,value,subkey){
 const row=billBatchResults[index];if(!row)return;
 const raw=String(value).trim();
 const v=raw.replace(/\\./g,'').replace(',', '.').replace(/[^0-9.\\-]/g,'');
 if(key==='tou_kwh'){
  row.tou_kwh=row.tou_kwh||{};
  const num=v===''?null:Number(v);
  if(v===''||!Number.isFinite(num)||num<0)delete row.tou_kwh[subkey];
  else row.tou_kwh[subkey]=num;
 }else{
  const num=v===''?null:Number(v);
  if(v===''||!Number.isFinite(num)||num<0)delete row[key];
  else row[key]=num;
 }
 if(row.monthly_kwh&&row.amount_vnd)row.average_price_vnd_kwh=row.amount_vnd/row.monthly_kwh;
 else delete row.average_price_vnd_kwh;
 row._manualEdited=true;
 const selected=$('billGroupSelect').value;
 renderBillBatch();
 if(selected&&[...$('billGroupSelect').options].some(o=>o.value===selected)){$('billGroupSelect').value=selected;applyBillGroup()}
 $('billReadStatus').textContent='Đã cập nhật số liệu nhập tay. Kiểm tra tổng khung giờ có khớp tổng kWh trước khi tính sizing.';
}
function billEditInput(index,key,value,step='1',subkey){
 const v=value===undefined||value===null?'':billFmt(value,{maximumFractionDigits:step==='1'?0:2});
 return '<input type="text" inputmode="decimal" value="'+esc(v)+'" aria-label="Nhập tay '+key+'" onchange="updateBillManual('+index+',&quot;'+key+'&quot;,this.value'+(subkey?',&quot;'+subkey+'&quot;':'')+')" style="width:96px;min-width:72px;padding:5px;border:1px solid #9bb9aa;border-radius:5px;background:#fff">';
}
function renderBillBatch(){
 const rows=billBatchResults;const seen=new Map();
 rows.forEach((x,i)=>{x._included=!!(x.monthly_kwh&&!x.error);x._duplicate=false;const k=billPeriodKey(x);if(!k||!x._included)return;const dedupeKey=billGroupKey(x)+'|'+k;if(seen.has(dedupeKey)){const j=seen.get(dedupeKey),a=rows[j],ad=Object.keys(a.tou_kwh||{}).length,bd=Object.keys(x.tou_kwh||{}).length;if(bd>ad){a._included=false;a._duplicate=true;seen.set(dedupeKey,i)}else{x._included=false;x._duplicate=true}}else seen.set(dedupeKey,i)});
 const included=rows.filter(x=>x._included),sum=(arr,key)=>arr.reduce((a,x)=>a+(Number(x[key])||0),0);
 const groups={};included.forEach(x=>{const k=billGroupKey(x);(groups[k]??=[]).push(x)});
 const groupKeys=Object.keys(groups);
 $('billGroupSelect').innerHTML='<option value="">Chọn nhóm hóa đơn</option>'+groupKeys.map(k=>'<option value="'+esc(k)+'">'+esc(k)+'</option>').join('');
 let summary='<b>Đã đọc '+rows.length+' file · '+included.length+' hóa đơn được tính · '+rows.filter(x=>x._duplicate).length+' file trùng cùng nhóm/kỳ bị loại.</b>';
 if(!included.length)summary+='<br>Chưa có hóa đơn nào đủ dữ liệu để tổng hợp.';
 $('billBatchSummary').innerHTML=summary;
 $('billGroupSummary').innerHTML=groupKeys.map(k=>{const arr=groups[k],kwh=sum(arr,'monthly_kwh'),money=sum(arr,'amount_vnd'),avg=kwh&&money?money/kwh:null,hasTou=arr.some(x=>Object.keys(x.tou_kwh||{}).length===3),normal=sum(arr.map(x=>({v:(x.tou_kwh||{}).normal})),'v'),peak=sum(arr.map(x=>({v:(x.tou_kwh||{}).peak})),'v'),off=sum(arr.map(x=>({v:(x.tou_kwh||{}).offpeak})),'v');return '<div style="padding:10px;margin:8px 0;border:1px solid #dce5e1;border-radius:8px"><b>'+esc(k)+'</b><br>'+arr.length+' hóa đơn · '+billFmt(kwh)+' kWh · Giá bình quân '+(avg?billFmt(avg)+' đ/kWh':'chưa đủ dữ liệu tiền điện')+'<br><small>'+(hasTou?'Bình thường '+billFmt(normal)+' · Cao điểm '+billFmt(peak)+' · Thấp điểm '+billFmt(off)+' kWh':'Chưa có đủ dữ liệu ba khung giờ')+'</small></div>'}).join('');
 $('billBatchTable').style.display=rows.length?'table':'none';
 $('billBatchRows').innerHTML=rows.map((x,i)=>'<tr style="border-top:1px solid #dce5e1"><td>'+esc(x.filename||'')+'<br><small>'+esc(x.period||'')+'</small></td><td>'+esc(x.customer_type||'Chưa rõ')+'</td><td>'+esc(x.tariff_type||'—')+'</td><td>'+esc(Object.entries(x.tiers||{}).map(([k,v])=>k.replace('bac_','Bậc ')+' '+billFmt(v.kwh)+' kWh').join('; ')||'—')+'</td><td>'+esc(x.voltage||'—')+'</td><td>'+billEditInput(i,'tou_kwh',(x.tou_kwh||{}).normal,'1','normal')+'</td><td>'+billEditInput(i,'tou_kwh',(x.tou_kwh||{}).peak,'1','peak')+'</td><td>'+billEditInput(i,'tou_kwh',(x.tou_kwh||{}).offpeak,'1','offpeak')+'</td><td>'+billEditInput(i,'monthly_kwh',x.monthly_kwh,'0.01')+'</td><td>'+billEditInput(i,'period_days',x.period_days)+'</td><td>'+billEditInput(i,'amount_vnd',x.amount_vnd,'1')+'</td><td>'+(x.average_price_vnd_kwh?billFmt(x.average_price_vnd_kwh):'—')+'</td><td>'+(x._duplicate?'Trùng kỳ — không cộng':x.error?esc(x.error):x._included?'Được tính':'Không tính')+(x._manualEdited?' · Đã sửa tay':'')+'</td></tr>').join('');
 if(groupKeys.length===1){$('billGroupSelect').value=groupKeys[0];applyBillGroup()}
 else if(groupKeys.length>1){$('billMonthKwh').value='';$('billDays').value=30;$('billDailyAverage').value='';$('billMonthlyAverageKwh').value='';$('billMonthlyAverageMoney').value='';$('billAveragePrice').value='';loadSizingResult=null;$('billReadStatus').textContent='Có nhiều nhóm loại điện/cấp áp. Chọn đúng nhóm trước khi tính sizing.'}
}

let billOcrWorker=null;
function loadBillScript(src){
 return new Promise((resolve,reject)=>{
  const existing=[...document.scripts].find(s=>s.src===src);
  if(existing){if(existing.dataset.loaded==='1')return resolve();existing.addEventListener('load',resolve,{once:true});existing.addEventListener('error',reject,{once:true});return}
  const s=document.createElement('script');s.src=src;s.onload=()=>{s.dataset.loaded='1';resolve()};s.onerror=()=>reject(Error('Không tải được thư viện OCR từ CDN. Kiểm tra kết nối Internet rồi thử lại.'));document.head.appendChild(s);
 });
}
async function getBillOcrWorker(){
 if(billOcrWorker)return billOcrWorker;
 await loadBillScript('https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js');
 if(!window.Tesseract)throw Error('Không khởi tạo được Tesseract.js.');
 billOcrWorker=await Tesseract.createWorker('vie',1,{
  logger:m=>{if(m.status==='recognizing text')$('billReadStatus').textContent='OCR trên máy tính: '+Math.round((m.progress||0)*100)+'%'},
  workerPath:'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
  corePath:'https://cdn.jsdelivr.net/npm/tesseract.js-core@5/tesseract-core.wasm.js',
  langPath:'https://tessdata.projectnaptha.com/4.0.0'
 });
 await billOcrWorker.setParameters({preserve_interword_spaces:'1'});
 return billOcrWorker;
}
async function billFileToCanvas(file){
 if(file.type==='application/pdf'||file.name.toLowerCase().endsWith('.pdf')){
  await loadBillScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
  if(!window.pdfjsLib)throw Error('Không tải được bộ đọc PDF trên trình duyệt.');
  pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
  const page=await pdf.getPage(1),viewport=page.getViewport({scale:1.7});
  const canvas=document.createElement('canvas');canvas.width=Math.round(viewport.width);canvas.height=Math.round(viewport.height);
  await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise;
  return canvas;
 }
 const url=URL.createObjectURL(file);
 try{
  const img=await new Promise((resolve,reject)=>{const im=new Image();im.onload=()=>resolve(im);im.onerror=()=>reject(Error('Không mở được ảnh hóa đơn.'));im.src=url});
  const scale=Math.min(1,2600/Math.max(img.naturalWidth,img.naturalHeight));
  const canvas=document.createElement('canvas');canvas.width=Math.round(img.naturalWidth*scale);canvas.height=Math.round(img.naturalHeight*scale);
  canvas.getContext('2d').drawImage(img,0,0,canvas.width,canvas.height);return canvas;
 }finally{URL.revokeObjectURL(url)}
}
async function ocrBillCanvas(canvas,worker){
 const w=canvas.width,h=canvas.height,outputs=[];
 // Read the full invoice first, then focused regions to preserve small figures.
 await worker.setParameters({tessedit_pageseg_mode:'11',preserve_interword_spaces:'1'});
 const full=await worker.recognize(canvas);
 outputs.push('=== OCR TOÀN TRANG ===\\n'+(full.data.text||''));
 await worker.setParameters({tessedit_pageseg_mode:'6',preserve_interword_spaces:'1'});
 const boxes=[[.02,.02,.98,.48],[.02,.28,.98,.78],[.02,.58,.98,.99]];
 for(let i=0;i<boxes.length;i++){
  const [x1,y1,x2,y2]=boxes[i],crop=document.createElement('canvas');
  crop.width=Math.max(1,Math.round((x2-x1)*w));crop.height=Math.max(1,Math.round((y2-y1)*h));
  crop.getContext('2d').drawImage(canvas,Math.round(x1*w),Math.round(y1*h),crop.width,crop.height,0,0,crop.width,crop.height);
  const result=await worker.recognize(crop,{rotateAuto:true});
  outputs.push('=== VÙNG '+(i+1)+' ===\\n'+(result.data.text||''));
 }
 return outputs.join('\\n');
}
async function ocrBillInBrowser(file){
 const worker=await getBillOcrWorker(),canvas=await billFileToCanvas(file);
 return ocrBillCanvas(canvas,worker);
}
async function readBillPdfInBrowser(file){
 await loadBillScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
 if(!window.pdfjsLib)throw Error('Không tải được bộ đọc PDF trên trình duyệt.');
 pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
 const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
 const nativePages=[];
 for(let n=1;n<=Math.min(pdf.numPages,8);n++){
  const page=await pdf.getPage(n),tc=await page.getTextContent(),rows=[];
  for(const item of tc.items||[]){
   const value=String(item.str||'').trim();if(!value)continue;
   const y=Number(item.transform?.[5]||0),x=Number(item.transform?.[4]||0);
   let row=rows.find(r=>Math.abs(r.y-y)<2.5);
   if(!row){row={y,items:[]};rows.push(row)}
   row.items.push({x,value});
  }
  const pageText=rows.sort((a,b)=>b.y-a.y).map(r=>r.items.sort((a,b)=>a.x-b.x).map(v=>v.value).join(' ')).join('\\n');
  nativePages.push('=== TRANG PDF '+n+' ===\\n'+pageText);
 }
 const nativeText=nativePages.join('\\n').slice(0,45000);
 let response=null,parsed={};
 if(nativeText.trim()){
  response=await fetch('/api/bills/parse-text',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename:file.name,text:nativeText})});
  try{parsed=await response.json()}catch(_){}
  if(!response.ok)throw Error(parsed.detail||'Không phân tích được nội dung PDF.');
 }
 // A PDF can contain selectable text but still omit figures in embedded image blocks.
 if(parsed.monthly_kwh&&parsed.amount_vnd){parsed.ocr_used=false;parsed.ocr_engine='pdfjs-text';return parsed}
 const worker=await getBillOcrWorker(),ocrPages=[];
 for(let n=1;n<=Math.min(pdf.numPages,3);n++){
  $('billReadStatus').textContent='PDF có lớp chữ chưa đủ dữ liệu; đang OCR trang '+n+'/'+Math.min(pdf.numPages,3)+' trên trình duyệt...';
  const page=await pdf.getPage(n),viewport=page.getViewport({scale:2.0});
  const canvas=document.createElement('canvas');canvas.width=Math.round(viewport.width);canvas.height=Math.round(viewport.height);
  await page.render({canvasContext:canvas.getContext('2d',{willReadFrequently:true}),viewport}).promise;
  ocrPages.push('=== OCR PDF TRANG '+n+' ===\\n'+await ocrBillCanvas(canvas,worker));
  const combined=(nativeText+'\\n'+ocrPages.join('\\n')).slice(0,49000);
  const check=await fetch('/api/bills/parse-text',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename:file.name,text:combined})});
  try{parsed=await check.json()}catch(_){}
  if(check.ok&&parsed.monthly_kwh&&parsed.amount_vnd)break;
 }
 const finalText=(nativeText+'\\n'+ocrPages.join('\\n')).slice(0,49000);
 response=await fetch('/api/bills/parse-text',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename:file.name,text:finalText})});
 try{parsed=await response.json()}catch(_){}
 if(!response.ok)throw Error(parsed.detail||'Không phân tích được PDF sau OCR.');
 parsed.ocr_used=ocrPages.length>0;parsed.ocr_engine=ocrPages.length?'pdfjs+tesseract-browser':'pdfjs-text';
 return parsed;
}
async function readBills(){
 const files=[...$('billFiles').files];
 if(!files.length){$('billReadStatus').textContent='Chọn ít nhất một hóa đơn trước.';return}
 if(files.length>50){$('billReadStatus').textContent='Mỗi lần tải tối đa 50 file.';return}
 billBatchResults=[];const progress=$('billReadProgress');
 const escHtml=s=>String(s||'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
 const renderProgress=(done,current,status)=>{
  const ok=billBatchResults.filter(x=>x.monthly_kwh&&!x.error).length,failed=billBatchResults.filter(x=>x.error).length;
  $('billReadStatus').textContent='Đọc hóa đơn '+done+'/'+files.length+' file · nhận diện được '+ok+'/'+done+' · lỗi '+failed;
  progress.innerHTML='<div style="height:8px;background:#e7efeb;border-radius:8px;overflow:hidden;margin:8px 0"><div style="height:100%;width:'+(done/files.length*100)+'%;background:#16845b\"></div></div><div style="font-size:13px;font-weight:600;margin-bottom:6px">'+escHtml(status||current)+'</div><div style="display:grid;gap:5px">'+files.map((file,i)=>{const item=billBatchResults[i];let label=item?(item.error?'Không đọc được: '+item.error:(item.monthly_kwh?'Đọc thành công · '+Number(item.monthly_kwh).toLocaleString('vi-VN')+' kWh':'Thiếu tổng kWh')):(i===done?'Đang OCR trên máy tính...':'Đang chờ');return '<div style="display:flex;justify-content:space-between;gap:12px;border-bottom:1px solid #edf1ef;padding:5px 0;font-size:12px"><span>'+ (i+1)+'. '+escHtml(file.name)+'</span><span>'+escHtml(label)+'</span></div>'}).join('')+'</div>';
 };
 renderProgress(0,files[0]?.name,'Chuẩn bị đọc PDF bằng trích xuất văn bản; ảnh sẽ dùng OCR trên Chrome.');
 try{
  if(files.some(f=>!(f.type==='application/pdf'||f.name.toLowerCase().endsWith('.pdf'))))await getBillOcrWorker();
  for(let i=0;i<files.length;i++){
   const file=files[i];renderProgress(i,file.name,'Đang '+((file.type==='application/pdf'||file.name.toLowerCase().endsWith('.pdf'))?'trích xuất PDF ':'OCR ảnh ')+(i+1)+'/'+files.length+': '+file.name);
   try{
    if(file.size>20*1024*1024)throw Error('File vượt quá 20 MB');
    let data={};
    if(file.type==='application/pdf'||file.name.toLowerCase().endsWith('.pdf')){
     // Read the PDF text layer in the browser; scanned/mixed PDFs fall back to browser OCR.
     data=await readBillPdfInBrowser(file);
    }else{
     // Raster images continue to use browser OCR.
     const text=await ocrBillInBrowser(file);
     if(!text.trim())throw Error('OCR không nhận diện được chữ trong file');
     const res=await fetch('/api/bills/parse-text',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename:file.name,text})});
     try{data=await res.json()}catch(_){}
     if(!res.ok)throw Error(data.detail||'Không phân tích được văn bản OCR');
    }
    billBatchResults.push(data);
   }catch(e){billBatchResults.push({filename:file.name,error:e.message,monthly_kwh:null})}
   renderBillBatch();calcBillNightShare();renderProgress(i+1,files[i+1]?.name,i+1===files.length?'Đã hoàn tất đọc hóa đơn.':'Đã xong '+(i+1)+'/'+files.length+' file.');
  }
 }catch(e){
  $('billReadStatus').textContent='Không khởi động được OCR trên Chrome: '+e.message;
  progress.innerHTML='<span style="color:#b42318">'+escHtml(e.message)+'</span>';return;
 }
 const ok=billBatchResults.filter(x=>x.monthly_kwh&&!x.error).length,fail=billBatchResults.length-ok;
 $('billReadStatus').textContent='Hoàn tất: đọc được '+ok+'/'+files.length+' file; '+fail+' file lỗi hoặc chưa nhận diện đủ tổng kWh.';
 $('loadSizingStatus').textContent='PDF ưu tiên đọc lớp chữ, tự OCR trên trình duyệt nếu thiếu dữ liệu; ảnh được OCR trên trình duyệt. Kiểm tra số liệu từng hóa đơn trước khi tính công suất.';
}
$('billDayShare').addEventListener('input',()=>{calcBillNightShare();loadSizingResult=null});
for(const id of ['billMonthKwh','billDays','billMonthlyAverageKwh','billMonthlyAverageMoney','billPsh','billDod'])$(id).addEventListener('input',()=>{loadSizingResult=null});

async function api(u,opt={}){let r=await fetch(u,opt);if(r.status===401){location='/login-page';throw Error('Cần đăng nhập lại')}let data={};try{data=await r.json()}catch(_){data={detail:'Máy chủ trả về lỗi không đọc được (HTTP '+r.status+').'}}if(!r.ok)throw Error(data.detail||data.message||'Lỗi HTTP '+r.status);return data}
async function save(){try{let p=payload();p.folderId=currentFolderId;let r=pid?await api('/api/projects/'+pid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)}):await api('/api/projects',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});pid=r.id;$('status').textContent='Đã lưu dự án vào thư mục.';list()}catch(e){$('status').textContent=e.message}}
async function list(){if(!currentFolderId)return;let a=await api('/api/folders/'+currentFolderId+'/projects');let html=a.map(x=>`<div class='project'><span><b>${esc(x.name)}</b><br><small>${esc(x.customer||'')}</small></span><span><button class='btn alt' onclick="openP('${x.id}')">Mở</button> <button class='btn alt' onclick="delP('${x.id}')">Xóa</button></span></div>`).join('')||'<div class="muted">Thư mục chưa có dự án. Nhấn Bắt đầu tạo dự án để thêm hồ sơ đầu tiên.</div>';$('projects').innerHTML=html;$('folderProjectList').innerHTML=html}
async function renderFolders(){let a=await api('/api/folders'),groups={};a.forEach(x=>{let d=x.date||'Chưa phân loại ngày';(groups[d]??=[]).push(x)});$('folderGroups').innerHTML=Object.keys(groups).sort((a,b)=>b.localeCompare(a)).map(d=>`<section class="folder-date-group"><div class="folder-date-title">${esc(d==='Chưa phân loại ngày'?d:new Date(d+'T00:00:00').toLocaleDateString('vi-VN',{weekday:'long',day:'2-digit',month:'2-digit',year:'numeric'}))}</div><div class="folder-grid">${groups[d].map(x=>`<div class="folder-card"><div class="folder-icon">📁</div><h3>${esc(x.name)}</h3><div class="muted">Thư mục dự án · ${esc(x.date||'')}</div><div class="folder-actions"><button class="btn" onclick="openFolder('${x.id}')">Mở thư mục</button><button class="btn alt" onclick="delFolder('${x.id}')">Xóa</button></div></div>`).join('')}</div></section>`).join('')||'<div class="panel muted">Chưa có thư mục. Hãy tạo thư mục đầu tiên để bắt đầu.</div>'}
async function createFolder(){let name=$('newFolderName').value.trim(),date=$('newFolderDate').value,btn=document.querySelector('.folder-create-row button');if(!name){$('folderStatus').textContent='Vui lòng nhập tên dự án/thư mục.';return}if(btn){btn.disabled=true;btn.textContent='ĐANG TẠO...'}$('folderStatus').textContent='Đang tạo thư mục và lưu dữ liệu...';try{let r=await api('/api/folders',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,date})});$('newFolderName').value='';await renderFolders();$('folderStatus').textContent='Đã tạo thư mục thành công.';await openFolder(r.id)}catch(e){$('folderStatus').textContent='Không tạo được: '+e.message;console.error('createFolder failed',e)}finally{if(btn){btn.disabled=false;btn.textContent='Tạo thư mục'}}}
async function openFolder(id){let a=await api('/api/folders'),f=a.find(x=>x.id===id);if(!f)return;currentFolderId=id;currentFolderName=f.name;currentFolderDate=f.date||'';$('folderViewTitle').textContent='📁 '+f.name;$('folderViewDate').textContent=f.date||'';$('dashboardView').style.display='none';$('editorView').style.display='none';$('folderView').style.display='block';await list()}
function startProject(){pid=null;fresh();$('currentFolderTitle').textContent=$('folderViewTitle').textContent;$('currentFolderDate').textContent=$('folderViewDate').textContent;$('folderView').style.display='none';$('editorView').style.display='block';$('status').textContent='Bắt đầu tạo dự án mới.'}
function backToFolder(){if(pid&&!confirm('Bạn đã lưu các thay đổi chưa? Quay lại thư mục?'))return;$('editorView').style.display='none';$('folderView').style.display='block';pid=null;list()}
function backToFolders(){if(!confirm('Quay lại danh sách tất cả thư mục?'))return;$('folderView').style.display='none';$('editorView').style.display='none';$('dashboardView').style.display='block';currentFolderId=null;pid=null;renderFolders()}
async function delFolder(id){if(!confirm('Xóa thư mục và toàn bộ dự án, file bên trong? Thao tác này không thể hoàn tác.'))return;try{await api('/api/folders/'+id,{method:'DELETE'});await renderFolders()}catch(e){$('folderStatus').textContent=e.message}}
function esc(s){return String(s||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
async function openP(id){let x=await api('/api/projects/'+id);pid=id;fill(x.payload);$('folderView').style.display='none';$('dashboardView').style.display='none';$('editorView').style.display='block';$('currentFolderTitle').textContent='📁 '+currentFolderName;$('currentFolderDate').textContent=currentFolderDate;let files=await listFiles();gsaImages=files.filter(v=>String(v.name||'').startsWith('global-solar-atlas-')).sort((a,b)=>a.name.localeCompare(b.name)).map(v=>({id:v.id,name:v.name}));renderGSAImages()}
async function delP(id){if(confirm('Xóa dự án?')){await api('/api/projects/'+id,{method:'DELETE'});if(pid===id)fresh();list()}}
function fresh(){pid=null;loadSizingResult=null;clearLoadProfile();$('billFiles').value='';billBatchResults=[];$('billBatchSummary').textContent='Chưa có dữ liệu hóa đơn.';$('billGroupSummary').innerHTML='';$('billGroupSelect').innerHTML='<option value="">Chọn nhóm hóa đơn</option>';$('billBatchTable').style.display='none';$('billMonthKwh').value='';$('billDays').value=30;$('billDayShare').value=65;calcBillNightShare();$('loadSizingResults').style.display='none';$('billReadStatus').textContent='';$('loadSizingStatus').textContent='Chưa tính toán.';financialOverrides={};showFinancialOverrides();$('financialStatus').textContent='';ids.forEach(k=>$(k).value=$(k).defaultValue||'');$('date').value=new Date().toISOString().slice(0,10);$('panelW').value=625;$('panelQty').value=30;$('batQty').value=0;$('vat').value=8;pvwattsResult=null;globalSolarAtlasResult=null;gsaImages=[];renderGSAImages();calc()}
async function upload(){if(!pid){alert('Hãy lưu dự án trước');return}let fd=new FormData();for(let f of $('files').files)fd.append('files',f);let r=await fetch('/api/projects/'+pid+'/files',{method:'POST',body:fd});if(!r.ok){alert('Tải file lỗi');return}$('files').value='';listFiles()}
async function uploadProjectImages(category){if(!pid){alert('Hãy lưu dự án trước khi tải ảnh.');return}let input=category==='roof-real'?$ ('roofRealImages'):$('roofPvImages');let files=[...input.files];if(!files.length){alert('Vui lòng chọn ảnh trước.');return}let fd=new FormData();for(let f of files)fd.append('files',f);fd.append('category',category);try{let r=await fetch('/api/projects/'+pid+'/files',{method:'POST',body:fd});let d={};try{d=await r.json()}catch(_){}if(!r.ok)throw Error(d.detail||'Tải ảnh không thành công');input.value='';await listFiles();$('status').textContent='Đã tải '+d.length+' ảnh lên dự án.'}catch(e){alert(e.message)}}
async function deleteProjectImage(fid){if(!pid||!fid)return;if(!confirm('Bạn có chắc muốn xóa ảnh này khỏi dự án?'))return;try{const r=await fetch('/api/projects/'+pid+'/files/'+encodeURIComponent(fid),{method:'DELETE'});let d={};try{d=await r.json()}catch(_){}if(!r.ok)throw Error(d.detail||'Không xóa được ảnh');await listFiles();$('status').textContent='Đã xóa ảnh khỏi dự án.'}catch(e){alert(e.message)}}
async function renderProjectImageLists(items){let real=items.filter(x=>String(x.name||'').startsWith('roof-real__')),pv=items.filter(x=>String(x.name||'').startsWith('roof-pv__'));const render=(a)=>a.length?a.map(x=>'<div style="display:flex;align-items:center;gap:8px;margin:7px 0"><img src="/api/files/'+x.id+'" style="width:76px;height:54px;object-fit:cover;border-radius:6px;border:1px solid #dce5e1"><a href="/api/files/'+x.id+'" target="_blank">'+esc(x.name.replace(/^roof-(real|pv)__/,''))+'</a><button class="btn alt" type="button" style="padding:5px 10px;margin-left:6px" onclick="deleteProjectImage(&quot;'+x.id+'&quot;)">Xóa</button></div>').join(''):'Chưa có ảnh.';$('roofRealImageList').innerHTML=render(real);$('roofPvImageList').innerHTML=render(pv)}

async function listFiles(){if(!pid)return;let a=await api('/api/projects/'+pid+'/files');let ordinary=a.filter(x=>!String(x.name||'').startsWith('roof-real__')&&!String(x.name||'').startsWith('roof-pv__'));$('fileList').innerHTML=ordinary.map(x=>`<div><a href='/api/files/${x.id}'>${esc(x.name)}</a></div>`).join('')||'Chưa có file';renderProjectImageLists(a);return a}
function exportJson(){let b=new Blob([JSON.stringify(payload(),null,2)],{type:'application/json'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='du-an.json';a.click()}
async function logout(){await fetch('/logout',{method:'POST'});location='/login-page'}
window.__solarInitStage='before-event-bindings';
const createFolderBtn=$('createFolderBtn');if(createFolderBtn)createFolderBtn.addEventListener('click',e=>{e.preventDefault();createFolder().catch(err=>{$('folderStatus').textContent='Không tạo được: '+err.message;console.error('createFolder click failed',err)})});initLoadProfileInputs();window.__solarInitStage='before-touScheduleMode-listener';$('touScheduleMode').addEventListener('change',()=>{updateTouScheduleDescription();if($('billGroupSelect').value)buildLoadProfileFromTou()});for(const id of ['plantWorkDays','plantStart','plantEnd','plantLunchStart','plantLunchEnd','plantLunchOff','plantOutsideShare'])$(id).addEventListener('change',()=>{if($('billGroupSelect').value)buildLoadProfileFromTou()});window.__solarInitStage='before-updateTouScheduleDescription';updateTouScheduleDescription();window.__solarInitStage='after-updateTouScheduleDescription';$('date').value=new Date().toISOString().slice(0,10);$('newFolderDate').value=new Date().toISOString().slice(0,10);window.__solarInitStage='before-calc';calc();window.__solarInitStage='after-calc-before-renderFolders';renderFolders().catch(e=>$('folderStatus').textContent=e.message);
</script><div id="financialPdfModal" style="display:none;position:fixed;inset:0;z-index:9999;background:rgba(10,20,16,.78);flex-direction:column;padding:12px"><div style="display:flex;align-items:center;gap:10px;background:#fff;border-radius:10px 10px 0 0;padding:10px 14px"><button class="btn alt" type="button" onclick="closeFinancialPdfPreview()">✕ Đóng</button><a id="financialPdfDownload" class="btn" href="#" download style="text-decoration:none">⬇ Tải PDF</a><span style="margin-left:auto;font-size:13px;color:#52665d">Báo cáo PTTC – Trang 5 và 6</span></div><iframe id="financialPdfFrame" title="Xem trước báo cáo PTTC PDF" style="width:100%;height:calc(100vh - 70px);border:0;background:#fff;border-radius:0 0 10px 10px"></iframe></div></main></body></html>"""

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

@app.get("/logout-now")
def logout_now(request:Request):
    request.session.clear()
    return RedirectResponse("/login-page",status_code=303)

@app.get("/",response_class=HTMLResponse)
def home(request:Request):
    if request.session.get("user")!=USER: return RedirectResponse("/login-page",status_code=303)
    return PAGE


_bill_ocr_env_logged=False

def _bill_ocr_image(image):
    global _bill_ocr_env_logged
    # EVNSPC bill layout: OCR only the useful left-side blocks. The table
    # crop is enlarged so small kWh digits remain legible, while charts/QR
    # codes and the contact column are excluded.
    image=ImageOps.exif_transpose(image).convert("L")
    # Large phone photos make each 2x crop disproportionately expensive on Render.
    # Normalize dimensions before cropping while retaining enough detail for bill text.
    image.thumbnail((1600,1600),Image.Resampling.LANCZOS)
    w,h=image.size
    if not _bill_ocr_env_logged:
        try:
            langs=pytesseract.get_languages(config="")
            version=str(pytesseract.get_tesseract_version())
            print(f"Bill OCR engine: Tesseract {version}; languages={langs}",flush=True)
            if "vie" not in langs:
                raise RuntimeError("Thiếu bộ ngôn ngữ Tesseract tiếng Việt (vie).")
            _bill_ocr_env_logged=True
        except Exception as exc:
            print(f"Bill OCR environment check failed: {type(exc).__name__}: {exc}",flush=True)
            raise

    boxes=[
        (0.05,0.02,0.66,0.36),  # customer type, voltage and billing period
        (0.07,0.26,0.57,0.57), # three time-of-use rows and total kWh
        (0.04,0.65,0.65,0.999) # electricity charges and total payment
    ]
    parts=[]
    for index,(x1,y1,x2,y2) in enumerate(boxes):
        crop=image.crop((int(w*x1),int(h*y1),int(w*x2),int(h*y2)))
        # These EVNSPC bills are already high-resolution (about 1,000 px wide).
        # Do not 2x-upscale every crop: it multiplies Tesseract work on small Render instances
        # without improving recognition of the printed kWh figures.
        crop=ImageOps.autocontrast(crop)
        crop_started=time.perf_counter()
        try:
            part=pytesseract.image_to_string(
                crop,lang="vie",config="--oem 1 --psm 6",timeout=12
            )
        except RuntimeError as exc:
            if "timeout" not in str(exc).lower():
                raise
            # Retry once at a smaller size; continue to other regions if still slow.
            crop.thumbnail((700,520),Image.Resampling.LANCZOS)
            try:
                part=pytesseract.image_to_string(
                    crop,lang="vie",config="--oem 1 --psm 6",timeout=7
                )
            except RuntimeError as retry_exc:
                if "timeout" not in str(retry_exc).lower():
                    raise
                print(f"Bill OCR crop {index+1} timed out twice; continuing with other regions",flush=True)
                part=""
        print(f"Bill OCR crop {index+1}: seconds={round(time.perf_counter()-crop_started,2)}; chars={len(part.strip())}",flush=True)
        parts.append(part)

    text="\n".join(parts)
    print(f"Bill OCR crop result: size={w}x{h}; chars={len(text.strip())}; crop_chars={[len(p.strip()) for p in parts]}",flush=True)
    # Try the complete page with more than one segmentation mode. This handles
    # bills whose layout differs from the assumed EVNSPC crop coordinates.
    if len(text.strip())<100:
        page=image.copy()
        page.thumbnail((1200,1600),Image.Resampling.LANCZOS)
        attempts=[
            (page,"vie","--oem 1 --psm 11",7),
            (page,"vie+eng","--oem 1 --psm 6",7),
            (ImageOps.autocontrast(page).point(lambda p: 255 if p>165 else 0),
             "vie+eng","--oem 1 --psm 11",7)
        ]
        best=text
        for attempt_no,(candidate,language,config,seconds) in enumerate(attempts,1):
            try:
                fallback=pytesseract.image_to_string(
                    candidate,lang=language,config=config,timeout=seconds
                )
                print(f"Bill OCR full-page attempt {attempt_no}: chars={len(fallback.strip())}; lang={language}; config={config}",flush=True)
                if len(fallback.strip())>len(best.strip()):
                    best=fallback
                if len(best.strip())>=100:
                    break
            except RuntimeError as exc:
                if "timeout" not in str(exc).lower():
                    raise
                print(f"Bill OCR full-page attempt {attempt_no} timed out; continuing",flush=True)
        text=best
    return text

def _bill_read_document(data,ext):
    text=""; used_ocr=False
    if ext==".pdf":
        reader=PdfReader(io.BytesIO(data))
        # Extract native PDF text from all invoice pages first. A fixed character
        # threshold is unreliable: some valid EVN PDFs have short but useful text.
        text="\n".join((p.extract_text() or "") for p in reader.pages[:8])
        native=text.strip()
        native_result=_bill_extract_values(native) if native else {}
        native_has_bill_data=bool(
            native_result.get("monthly_kwh") or
            any((native_result.get("tou_kwh") or {}).values()) or
            native_result.get("amount_vnd")
        )
        # Keep native text whenever it contains usable bill values; OCR only
        # image-only PDFs or PDFs whose text layer contains no usable bill data.
        if not native_has_bill_data:
            page_count=min(3,len(reader.pages))
            pages=convert_from_bytes(data,dpi=240,first_page=1,last_page=page_count)
            ocr_text="\n".join(_bill_ocr_image(page) for page in pages)
            if ocr_text.strip():
                text=ocr_text
                used_ocr=True
    else:
        with Image.open(io.BytesIO(data)) as source:
            text=_bill_ocr_image(source)
        used_ocr=True
    return text,used_ocr

def _bill_number(raw):
    raw=(raw or "").strip().replace(" ","")
    if not raw: return None
    if "," in raw and "." in raw:
        if raw.rfind(",")>raw.rfind("."): raw=raw.replace(".","").replace(",",".")
        else: raw=raw.replace(",","")
    elif "," in raw:
        tail=raw.rsplit(",",1)[-1]
        raw=raw.replace(",","." if len(tail)<=2 else "")
    elif "." in raw:
        tail=raw.rsplit(".",1)[-1]
        if len(tail)==3: raw=raw.replace(".","")
    try: return float(raw)
    except ValueError: return None

def _bill_extract_values(text):
    normalized=re.sub(r"\s+"," ",text or "").strip()
    lines=[re.sub(r"\s+"," ",x).strip() for x in (text or "").splitlines() if x.strip()]
    # Read each OCR line independently. Flattening the whole page can make
    # a regex jump from a TOU label into the unrelated payment table below.
    import unicodedata
    def _bill_plain(value):
        return "".join(ch for ch in unicodedata.normalize("NFD",value or "")
                       if unicodedata.category(ch)!="Mn").replace("đ","d").replace("Đ","D").lower()
    # Parse only the meter-consumption section. Some EVNSPC bills replace
    # the meter mid-cycle and repeat each TOU row for the old and new meters.
    # Sum those rows; stop before the separate tariff/payment breakdown.
    tou={}
    in_meter_section=False
    pending_key=None
    for line in lines:
        plain_line=_bill_plain(line)
        if "cong to do dem" in plain_line:
            in_meter_section=True
            pending_key=None
            continue
        if not in_meter_section:
            continue
        if re.search(r"tong\s*:",plain_line):
            break
        key=None
        if "khung gio binh thuong" in plain_line: key="normal"
        elif "khung gio cao diem" in plain_line: key="peak"
        elif "khung gio thap diem" in plain_line: key="offpeak"
        if key:
            pending_key=key
        values=re.findall(r"\d[\d.,]*",line)
        # Meter rows contain multiplier, new index, old index, consumption.
        # Wrapped removal dates (e.g. 26/03/2026) are skipped until the
        # following numeric row supplies all four fields.
        if pending_key and len(values)>=4:
            value=_bill_number(values[3])
            if value is not None:
                tou[pending_key]=tou.get(pending_key,0)+value
                pending_key=None
    total_hint=None
    for line in lines:
        plain_line=_bill_plain(line)
        tm_total=re.search(r"tong\s*:\s*([\d.,]+)",plain_line,re.I)
        if tm_total:
            total_hint=_bill_number(tm_total.group(1))
            break
    # The OCR may misread one digit in a TOU row (for example 2.718 as 2.018).
    # When the invoice's explicit "Tổng:" line exists, it is the authoritative
    # monthly quantity; do not silently replace it with the sum of noisy rows.
    tou_sum=sum(tou.values()) if len(tou)==3 else None
    kwh=total_hint if total_hint and total_hint>0 else tou_sum
    if total_hint and len(tou)==2:
        missing=({"normal","peak","offpeak"}-set(tou))
        inferred=total_hint-sum(tou.values())
        if len(missing)==1 and inferred>=0:
            tou[next(iter(missing))]=inferred
    # If no explicit total is printed, fall back to the three recognized TOU values.
    if kwh is None and tou_sum is not None:
        kwh=tou_sum
    if kwh is None:
        # EVN VAT invoice form: one line item "Điện tiêu thụ kỳ ...", unit kWh,
        # quantity in the quantity column, then unit price and amount.
        invoice_patterns=[
            r"Điện tiêu thụ.{0,220}?\bkwh\s+([\d.,]+)\s+(?:-|[\d.,]+)\s+[\d.,]+",
            r"Điện tiêu thụ.{0,220}?\bkwh\s+([\d.,]+)"
        ]
        for pat in invoice_patterns:
            m=re.search(pat,normalized,re.I)
            if m:
                value=_bill_number(m.group(1))
                if value and value>0:
                    kwh=value
                    break
    if kwh is None:
        # Common EVNSPC total row: "Tổng: 19.247" following the consumption table.
        m=re.search(r"Tổng\s*:\s*([\d.,]+)",normalized,re.I)
        if m:
            value=_bill_number(m.group(1))
            if value and value>0: kwh=value
    if kwh is None:
        # Alternate EVN layouts print the total after the explicit energy-total heading.
        m=re.search(r"Tổng điện năng tiêu thụ\s*\(kWh\)\s*([\d.,]+)",normalized,re.I)
        if m:
            value=_bill_number(m.group(1))
            if value and value>0: kwh=value
    if kwh is None:
        patterns=[
            r"(?:điện năng tiêu thụ|sản lượng điện tiêu thụ|tổng điện năng|điện năng sử dụng|điện tiêu thụ|tổng sản lượng)[^\d]{0,60}([\d.,]+)\s*kwh",
            r"([\d.,]+)\s*kwh[^\n]{0,60}(?:điện năng tiêu thụ|sản lượng|tiêu thụ)"
        ]
        for line in lines:
            low=line.lower()
            if any(x in low for x in ("đơn giá","thành tiền","chỉ số công tơ","điện năng phản kháng")): continue
            for pat in patterns:
                m=re.search(pat,line,re.I)
                if m:
                    value=_bill_number(m.group(1))
                    if value and value>0:
                        kwh=value; break
            if kwh is not None: break
    period_days=None; period=""
    m=re.search(r"(?:từ ngày|từ)\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4}).{0,80}?(?:đến ngày|đến)\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4})",normalized,re.I)
    if m:
        period=m.group(1)+" đến "+m.group(2)
        try:
            a=datetime.strptime(m.group(1).replace("-","/"),"%d/%m/%Y")
            b=datetime.strptime(m.group(2).replace("-","/"),"%d/%m/%Y")
            period_days=max(1,(b-a).days+1)  # Billing periods include both start and end dates.
        except Exception: pass
    # Billing metadata: classify only when the wording is explicit.
    low=normalized.lower()
    if any(x in low for x in ("sinh hoạt","điện sinh hoạt","bậc 1","bậc 2")):
        customer_type="Sinh hoạt"
    elif any(x in low for x in ("kinh doanh","mục đích kinh doanh")):
        customer_type="Kinh doanh"
    elif any(x in low for x in ("sản xuất","mục đích sản xuất")):
        customer_type="Sản xuất"
    elif any(x in low for x in ("hành chính sự nghiệp","cơ quan hành chính","đơn vị sự nghiệp")):
        customer_type="Hành chính sự nghiệp"
    else:
        customer_type="Chưa xác định"
    vm=re.search(r"(?:cấp điện áp|điện áp sử dụng|điện áp)\s*[:\-]?\s*(dưới\s*6\s*kv|từ\s*6\s*kv\s*đến\s*dưới\s*22\s*kv|từ\s*22\s*kv\s*đến\s*dưới\s*35\s*kv|từ\s*35\s*kv\s*đến\s*dưới\s*110\s*kv|110\s*kv\s*trở\s*lên|[\d.,]+\s*kv)",normalized,re.I)
    voltage=vm.group(1) if vm else ""
    amount=None
    amount_patterns=[
        r"(?:tổng cộng tiền thanh toán|tổng cộng thanh toán|tổng tiền thanh toán|tổng thanh toán|số tiền thanh toán|cộng tiền thanh toán)[^\d]{0,50}([\d.,]+)\s*(?:đ|vnđ|vnd)?",
        r"(?:tổng tiền điện|tiền điện sau thuế|tổng tiền)[^\d]{0,50}([\d.,]+)\s*(?:đ|vnđ|vnd)"
    ]
    for pat in amount_patterns:
        mm=re.search(pat,normalized,re.I)
        if mm:
            amount=_bill_number(mm.group(1))
            # Reject OCR fragments such as "23,48" captured from a split
            # thousands-group; an EVN invoice payment total should be > 1 million VND.
            if amount and amount>=1000000: break
            amount=None
    if not amount:
        # Search the explicit payment-total line, not the nearby "Tổng số tiền"
        # heading, which can otherwise cause OCR to capture a tariff-table number.
        for line in lines:
            plain_line=_bill_plain(line)
            if "tong cong tien thanh toan" in plain_line or "tong tien thanh toan" in plain_line:
                values=re.findall(r"\d[\d.,]*",line)
                if values:
                    candidate=_bill_number(values[-1])
                    if candidate and candidate>1000000:
                        amount=candidate
                        break
    start_iso=end_iso=""
    dm=re.search(r"(?:từ ngày|từ)\s*(\d{1,2})[/-](\d{1,2})[/-](\d{4}).{0,100}?(?:đến ngày|đến)\s*(\d{1,2})[/-](\d{1,2})[/-](\d{4})",normalized,re.I)
    if dm:
        try:
            da=datetime(int(dm.group(3)),int(dm.group(2)),int(dm.group(1)))
            dbb=datetime(int(dm.group(6)),int(dm.group(5)),int(dm.group(4)))
            start_iso=da.strftime("%Y-%m-%d"); end_iso=dbb.strftime("%Y-%m-%d")
        except Exception: pass
    if not start_iso:
        dm=re.search(r"(?:kỳ|tháng)\s*(\d{1,2})\s*[/\-]\s*(\d{4})",normalized,re.I)
        if dm:
            try:
                da=datetime(int(dm.group(2)),int(dm.group(1)),1)
                start_iso=da.strftime("%Y-%m-%d")
            except Exception: pass
    # Residential tier rows are retained separately; never infer missing tier quantities.
    tiers={}
    for line in lines:
        tm=re.search(r"bậc\s*(\d+)\s*[^\d]{0,35}([\d.,]+)\s*(?:kwh)?\s+([\d.,]+)\s+([\d.,]+)",line,re.I)
        if tm:
            qty=_bill_number(tm.group(2))
            rate=_bill_number(tm.group(3))
            money=_bill_number(tm.group(4))
            if qty is not None:
                tiers["bac_"+tm.group(1)]={"kwh":qty,"unit_price":rate,"amount_vnd":money}
    tariff="Ba giá" if len(tou)==3 else ("Bậc thang" if customer_type=="Sinh hoạt" else ("Theo bậc thang" if tiers else ("Một giá / tổng hợp" if kwh is not None else "Chưa xác định")) )
    return {"monthly_kwh":kwh,"period_days":period_days,"period":period,"period_start":start_iso,"period_end":end_iso,
            "tou_kwh":tou,"tiers":tiers,"customer_type":customer_type,"voltage":voltage,"amount_vnd":amount,
            "average_price_vnd_kwh":(amount/kwh if amount and kwh else None),"tariff_type":tariff,
            "raw_text":"\n".join(lines)[:7000]}
@app.post("/api/bill/read")
async def read_electricity_bill(request:Request,file:UploadFile=File(...)):
    auth(request)
    name=Path(file.filename or "bill").name
    ext=Path(name).suffix.lower()
    if ext not in (".pdf",".png",".jpg",".jpeg",".webp"):
        raise HTTPException(400,"Chỉ hỗ trợ hóa đơn PDF, PNG, JPG hoặc WEBP.")
    data=await file.read()
    if not data or len(data)>20*1024*1024:
        raise HTTPException(413,"File hóa đơn trống hoặc vượt quá 20 MB.")
    text=""; used_ocr=False
    try:
        text,used_ocr=await asyncio.to_thread(_bill_read_document,data,ext)
    except Exception as exc:
        print(f"Bill parsing error: {type(exc).__name__}: {exc}",flush=True)
        detail=("OCR quá thời gian xử lý. Vui lòng thử lại ảnh rõ hơn hoặc nhập số liệu bằng tay."
                if isinstance(exc,RuntimeError) and "timeout" in str(exc).lower()
                else "Không đọc được file hóa đơn. Bạn có thể nhập số liệu bằng tay.")
        raise HTTPException(422,detail)
    if not text.strip():
        raise HTTPException(422,"Không nhận diện được nội dung hóa đơn. Vui lòng nhập số liệu bằng tay.")
    result=_bill_extract_values(text)
    result["filename"]=name
    result["ocr_used"]=used_ocr
    result["message"]="Đã trích xuất văn bản; vui lòng kiểm tra và xác nhận sản lượng, số ngày trước khi tính toán."
    return result

@app.post("/api/bills/parse-text")
async def parse_bill_ocr_text(request:Request):
    auth(request)
    try:
        payload=await request.json()
    except Exception:
        raise HTTPException(400,"Dữ liệu OCR không hợp lệ.")
    filename=Path(str(payload.get("filename") or "bill")).name
    text=str(payload.get("text") or "")
    if not text.strip():
        raise HTTPException(422,"Không có nội dung OCR để phân tích.")
    if len(text)>50000:
        raise HTTPException(413,"Nội dung OCR vượt quá giới hạn.")
    result=_bill_extract_values(text)
    result["filename"]=filename
    result["ocr_used"]=True
    result["ocr_engine"]="browser"
    result["error"]=None if result.get("monthly_kwh") else "Chưa nhận diện được tổng kWh"
    result["message"]="Đã nhận văn bản OCR từ trình duyệt; vui lòng kiểm tra số liệu."
    return result

@app.post("/api/bills/read")
async def read_electricity_bills(request:Request,files:List[UploadFile]=File(...)):
    auth(request)
    if not files: raise HTTPException(400,"Vui lòng chọn ít nhất một hóa đơn.")
    if len(files)>50: raise HTTPException(400,"Mỗi lần tải tối đa 50 file.")
    results=[]
    for file in files:
        name=Path(file.filename or "bill").name
        ext=Path(name).suffix.lower()
        if ext not in (".pdf",".png",".jpg",".jpeg",".webp"):
            results.append({"filename":name,"error":"Định dạng không hỗ trợ","monthly_kwh":None}); continue
        data=await file.read()
        if not data or len(data)>20*1024*1024:
            results.append({"filename":name,"error":"File trống hoặc vượt quá 20 MB","monthly_kwh":None}); continue
        text=""; used_ocr=False
        started=time.perf_counter()
        try:
            text,used_ocr=await asyncio.to_thread(_bill_read_document,data,ext)
            elapsed=round(time.perf_counter()-started,2)
            print(f"Bill OCR completed: {name}; seconds={elapsed}; ocr={used_ocr}",flush=True)
            if not text.strip():
                results.append({"filename":name,"error":"Không nhận diện được nội dung","monthly_kwh":None}); continue
            item=_bill_extract_values(text)
            item["filename"]=name; item["ocr_used"]=used_ocr
            item["error"]=None if item.get("monthly_kwh") else "Chưa nhận diện được tổng kWh"
            results.append(item)
        except Exception as exc:
            print(f"Batch bill parsing error {name}: {type(exc).__name__}: {exc}",flush=True)
            reason=("OCR quá thời gian xử lý; thử ảnh rõ hơn hoặc nhập tay"
                    if isinstance(exc,RuntimeError) and "timeout" in str(exc).lower()
                    else f"Lỗi OCR: {type(exc).__name__}")
            results.append({"filename":name,"error":reason,"monthly_kwh":None})
    return {"count":len(results),"results":results}

@app.post("/api/quote-parse")
async def parse_quote_pdf(request:Request,file:UploadFile=File(...)):
    auth(request)
    name=Path(file.filename or "").name
    if Path(name).suffix.lower()!=".pdf":
        raise HTTPException(400,"Vui lòng tải lên báo giá định dạng PDF.")
    data=await file.read()
    if not data or len(data)>20*1024*1024:
        raise HTTPException(413,"File PDF rỗng hoặc vượt quá giới hạn 20 MB.")
    try:
        from pypdf import PdfReader
        reader=PdfReader(io.BytesIO(data))
        text_content="\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        raise HTTPException(422,"Không thể đọc cấu trúc PDF: "+str(exc)[:200])
    method="text"
    normalized=text_content.replace("\u00a0"," ")
    def parse_amount_block(text):
        # Keep PDF text extraction as primary; OCR only when text is absent or totals are incomplete.
        def clean_label(s):
            import unicodedata
            s=s.replace("đ","d").replace("Đ","D")
            s=unicodedata.normalize("NFD",s)
            s="".join(ch for ch in s if unicodedata.category(ch)!="Mn")
            return re.sub(r"[^a-z0-9]+"," ",s.lower()).strip()
        def nums(s):
            found=[]
            for raw in re.findall(r"(?<!\d)(?:\d{1,3}(?:[., ]\d{3})+|\d{6,})(?!\d)",s):
                digits=re.sub(r"\D","",raw)
                if digits:
                    value=int(digits)
                    if value>=1000 and value not in found: found.append(value)
            return found
        lines=[re.sub(r"[ \t]+"," ",x).strip() for x in text.splitlines() if x.strip()]
        norm=[clean_label(x) for x in lines]
        def near(labels, reject=()):
            for i,line in enumerate(norm):
                if any(label in line for label in labels) and not any(x in line for x in reject):
                    # PDF extractors may place the amount before, after, or on the next line.
                    block=" ".join(lines[max(0,i-1):min(len(lines),i+3)])
                    candidates=nums(block)
                    if candidates:
                        # Prefer values in the label's own line; otherwise nearest following line.
                        own=nums(lines[i])
                        if own:return own[-1]
                        for j in range(i+1,min(len(lines),i+3)):
                            nxt=nums(lines[j])
                            if nxt:return nxt[-1]
                        return candidates[-1]
            return None
        total=near([
            "tong tien da bao gom vat","tong tien bao gom vat","tong cong bao gom vat",
            "tong thanh toan","tong cong thanh toan","tong tien thanh toan",
            "tong gia tri sau thue","tong sau vat","so tien phai thanh toan"
        ])
        subtotal=near([
            "tong tien chua bao gom vat","tong tien truoc vat","tong truoc vat",
            "gia tri truoc thue","cong tien hang","tong cong truoc thue"
        ])
        vat=near(["tien thue gtgt","thue gtgt","tien vat","thue vat","vat"])
        # Some quotations print only three consecutive totals in a summary block.
        if total is None or subtotal is None:
            for i,line in enumerate(norm):
                if ("chua bao gom vat" in line or "truoc vat" in line) and i>=1:
                    window=" ".join(lines[max(0,i-4):min(len(lines),i+4)])
                    candidates=nums(window)
                    for a in range(len(candidates)-2):
                        sub,tax,grand=candidates[a:a+3]
                        if sub>0 and grand>=sub and abs(sub+tax-grand)<=max(2,grand*0.003):
                            subtotal,vat,total=sub,tax,grand
        if total is not None and subtotal is not None and vat is None:
            diff=total-subtotal
            if diff>=0:vat=diff
        if total is not None and vat is not None and subtotal is None:
            diff=total-vat
            if diff>=0:subtotal=diff
        if total is None and subtotal is not None and vat is not None:
            total=subtotal+vat
        return subtotal,vat,total
    subtotal,vat,total=parse_amount_block(normalized)
    # Highest-confidence path: inspect each PDF page independently.
    # Normalize labels locally; clean_label is scoped inside parse_amount_block.
    def normalize_quote_label(value):
        import unicodedata
        value=value.replace("đ","d").replace("Đ","D")
        value=unicodedata.normalize("NFD",value)
        value="".join(ch for ch in value if unicodedata.category(ch)!="Mn")
        return re.sub(r"[^a-z0-9]+"," ",value.lower()).strip()
    page_triplets=[]
    for page_index,page in enumerate(reader.pages):
        page_text=(page.extract_text() or "").replace("\u00a0"," ")
        page_labels=normalize_quote_label(page_text)
        if "tong tien" not in page_labels or "vat" not in page_labels:
            continue
        page_tokens=re.findall(r"(?<!\d)(?:\d{1,3}(?:[,. ]\d{3})+|\d{6,})(?!\d)",page_text)
        page_values=[int(re.sub(r"\D","",v)) for v in page_tokens]
        for token_index in range(max(0,len(page_values)-2)):
            sub,tax,grand=page_values[token_index:token_index+3]
            if sub>=100000 and tax>0 and grand>sub and abs(sub+tax-grand)<=max(2,grand*0.00001):
                tax_rate=tax/sub
                if 0.01<=tax_rate<=0.20:
                    page_triplets.append((page_index,token_index,sub,tax,grand))
    if page_triplets:
        _,_,subtotal,vat,total=page_triplets[-1]
    # Many supplier PDFs extract the right-aligned amount column before the row labels.
    # Recover the three summary amounts by arithmetic reconciliation across the document.
    if subtotal is None or vat is None or total is None:
        money_tokens=re.findall(r"(?<!\d)(?:\d{1,3}(?:[,. ]\d{3})+|\d{6,})(?!\d)",normalized)
        money_values=[int(re.sub(r"\D","",v)) for v in money_tokens]
        candidates=[]
        for idx in range(max(0,len(money_values)-2)):
            sub,tax,grand=money_values[idx:idx+3]
            if sub>=100000 and tax>=0 and grand>=sub and abs(sub+tax-grand)<=max(2,grand*0.003):
                candidates.append((sub,tax,grand))
        if candidates:
            subtotal,vat,total=candidates[-1]
    # Dedicated fast quotation OCR. Never reuse the multi-crop EVN bill OCR pipeline here.
    # Moderate rendering resolution and hard per-page Tesseract time limits keep this request bounded.
    if subtotal is None or vat is None or total is None:
        try:
            page_count=min(3,len(reader.pages))
            pages=convert_from_bytes(data,dpi=150,first_page=1,last_page=page_count,fmt="jpeg")
            ocr_parts=[]
            for page_index,page in enumerate(pages,1):
                page=ImageOps.autocontrast(ImageOps.exif_transpose(page).convert("L"))
                page.thumbnail((1400,1800),Image.Resampling.LANCZOS)
                try:
                    page_text=pytesseract.image_to_string(page,lang="vie+eng",config="--oem 1 --psm 6",timeout=12)
                except RuntimeError as exc:
                    if "timeout" not in str(exc).lower(): raise
                    print(f"Quote OCR page {page_index} timed out; trying sparse layout once",flush=True)
                    try:
                        page_text=pytesseract.image_to_string(page,lang="vie+eng",config="--oem 1 --psm 11",timeout=8)
                    except RuntimeError as retry_exc:
                        if "timeout" not in str(retry_exc).lower(): raise
                        page_text=""
                ocr_parts.append(page_text)
                ocr_sub,ocr_vat,ocr_total=parse_amount_block(page_text)
                subtotal=subtotal if subtotal is not None else ocr_sub
                vat=vat if vat is not None else ocr_vat
                total=total if total is not None else ocr_total
                print(f"Quote OCR page {page_index}: chars={len(page_text.strip())}; totals={subtotal}/{vat}/{total}",flush=True)
                if subtotal is not None and vat is not None and total is not None:
                    break
            method="text+ocr" if normalized.strip() else "ocr"
        except Exception as exc:
            print(f"Quote PDF OCR fallback failed: {type(exc).__name__}: {exc}",flush=True)
    if total is None and subtotal is not None and vat is not None:
        total=subtotal+vat
    if subtotal is None and total is not None and vat is not None:
        subtotal=total-vat if total>=vat else None
    if vat is None and total is not None and subtotal is not None:
        vat=total-subtotal if total>=subtotal else None
    if total is None:
        raise HTTPException(422,"Chưa nhận diện được tổng tiền trên PDF. Hãy kiểm tra trang tổng hợp báo giá có rõ dòng trước VAT, VAT và sau VAT.")
    if subtotal is not None and vat is not None and abs((subtotal+vat)-total)>max(2,total*0.003):
        print(f"Quote totals mismatch: subtotal={subtotal}, vat={vat}, total={total}",flush=True)
    # Read equipment values from their own quotation sections to avoid picking up unrelated numbers.
    def section(start_pattern,end_pattern=None):
        m=re.search(start_pattern,normalized,re.I)
        if not m:return ""
        tail=normalized[m.end():]
        if end_pattern:
            e=re.search(end_pattern,tail,re.I)
            if e:tail=tail[:e.start()]
        return tail
    panel_section=section(r"1\s*TẤM PIN NLMT",r"2\s*INVERTER")
    inverter_section=section(r"2\s*INVERTER",r"3\s*PIN LƯU TRỮ")
    battery_section=section(r"3\s*PIN LƯU TRỮ",r"4\s*HỆ THỐNG")
    def first(pattern,source,flags=re.I):
        m=re.search(pattern,source,flags)
        return m.group(1).strip(" :;,-") if m else None
    panel_brand=first(r"Thương hiệu\s*([^\n]+)",panel_section)
    panel_qty_raw=first(r"\bTấm\s+(\d{1,4})\b",panel_section)
    panel_qty=int(panel_qty_raw) if panel_qty_raw else None
    panel_wp_raw=first(r"(?:≥|>=)\s*(\d{3,4})",panel_section)
    if not panel_wp_raw:
        # The cover states total DC capacity; derive module rating only when panel count is explicit.
        headline=first(r"Công suất\s*(\d+(?:[.,]\d+)?)\s*kWp",normalized)
        panel_wp_raw=str(round(float(headline.replace(",","."))*1000/panel_qty)) if headline and panel_qty else None
    panel_wp=int(float(panel_wp_raw)) if panel_wp_raw else None
    inverter_model=first(r"(SOLIS[^\n]{0,100})",inverter_section)
    inverter_kw_raw=first(r"(?:HYBRID[^\n]{0,40}?|)(\d+(?:[.,]\d+)?)\s*KW",inverter_section)
    inverter_kw=float(inverter_kw_raw.replace(",",".")) if inverter_kw_raw else None
    phase_raw=first(r"(1\s*PHA|3\s*PHA|MỘT PHA|BA PHA)",inverter_section)
    phase=("1 pha" if phase_raw and ("1" in phase_raw or "MỘT" in phase_raw.upper()) else "3 pha" if phase_raw else None)
    battery_model=first(r"(Pylontech[^\n]*)",battery_section)
    battery_kwh_raw=first(r"(\d+(?:[.,]\d+)?)\s*kwh",battery_section)
    battery_kwh=float(battery_kwh_raw.replace(",",".")) if battery_kwh_raw else None
    battery_qty_raw=first(r"\bBộ\s+(\d{1,3})\b",battery_section)
    battery_qty=int(battery_qty_raw) if battery_qty_raw else None
    # Customer values can be emitted before table labels by some PDF generators.
    all_lines=[re.sub(r"[ \t]+"," ",line).strip() for line in normalized.splitlines() if line.strip()]
    customer_name=None
    site_address=None
    customer_phone=None
    # In this supplier template, the customer name and address are the first two
    # text lines on page 1, before the item table and supplier header.
    if all_lines:
        first=all_lines[0]
        if not re.search(r"^(Công ty|BÁO GIÁ|Địa chỉ:|Điện thoại:|Web:)",first,re.I):
            customer_name=first
    for line in all_lines[:6]:
        if re.search(r"(Kp\s*\d+|khu phố|phường|quận|huyện|thành phố|tỉnh|việt nam)",line,re.I):
            site_address=line
            break
    # Fallback to a conventional label/value layout when the PDF keeps values beside labels.
    def value_after_label(label):
        for ix,line in enumerate(all_lines):
            m=re.match(r"^"+label+r"\b\s*[:：]?\s*(.*)$",line,re.I)
            if m and m.group(1).strip():
                return m.group(1).strip()
            if m and ix+1<len(all_lines):
                nxt=all_lines[ix+1].strip()
                if nxt and not re.match(r"^(Mô tả chung|STT|Email|Điện thoại|BÁO GIÁ)",nxt,re.I):
                    return nxt
        return None
    if not customer_name:
        customer_name=value_after_label("Tên")
    if not site_address:
        site_address=value_after_label("Địa chỉ")
    # Restrict phone lookup to the customer block before the supplier quote heading;
    # never mistake the supplier's own phone number in its letterhead for the customer's.
    quote_heading=next((ix for ix,line in enumerate(all_lines) if re.search(r"BÁO GIÁ DỰ ÁN",line,re.I)),len(all_lines))
    customer_lines=all_lines[:quote_heading]
    for ix,line in enumerate(customer_lines):
        m=re.match(r"^Điện thoại\b\s*[:：]?\s*(.*)$",line,re.I)
        if m:
            candidate=m.group(1).strip()
            if not candidate and ix+1<len(customer_lines):
                candidate=customer_lines[ix+1].strip()
            digits=re.sub(r"\D","",candidate)
            if 9<=len(digits)<=11:
                customer_phone=digits
                break
    return {"filename":name,"customer_name":customer_name,"customer_phone":customer_phone,"site_address":site_address,"subtotal":subtotal,"vat":vat,"total_including_vat":total,"panel_brand":panel_brand,"panel_wp":panel_wp,"panel_qty":panel_qty,"inverter_model":inverter_model,"inverter_kw":inverter_kw,"phase":phase,"battery_model":battery_model,"battery_kwh":battery_kwh,"battery_qty":battery_qty,"method":method,"page_count":len(reader.pages)}


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
            "inv_eff":p["inv_eff"],"gcr":p["gcr"]
        }
        if int(p.get("use_wf_albedo",1))==1: params["use_wf_albedo"]=1
        elif p.get("albedo") is not None: params["albedo"]=p["albedo"]
        monthly_params={**params,"timeframe":"monthly"}
        resp=requests.get("https://developer.nlr.gov/api/pvwatts/v8.json",params=monthly_params,timeout=60)
        try: data=resp.json()
        except ValueError: raise HTTPException(502,"PVWatts trả về dữ liệu không hợp lệ.")
        if resp.status_code>=400 or data.get("errors"):
            detail="; ".join(data.get("errors") or [f"PVWatts HTTP {resp.status_code}"])
            raise HTTPException(422,detail[:600])
        out=data.get("outputs",{})
        monthly=out.get("ac_monthly",[])
        if len(monthly)!=12: raise HTTPException(502,"PVWatts không trả đủ 12 giá trị sản lượng tháng.")
        hourly_params={**params,"timeframe":"hourly"}
        hresp=requests.get("https://developer.nlr.gov/api/pvwatts/v8.json",params=hourly_params,timeout=90)
        try: hdata=hresp.json()
        except ValueError: raise HTTPException(502,"PVWatts hourly trả về dữ liệu không hợp lệ.")
        if hresp.status_code>=400 or hdata.get("errors"):
            detail="; ".join(hdata.get("errors") or [f"PVWatts hourly HTTP {hresp.status_code}"])
            raise HTTPException(422,detail[:600])
        ac=hdata.get("outputs",{}).get("ac",[])
        if len(ac)!=8760: raise HTTPException(502,f"PVWatts hourly trả về {len(ac)} giờ, không phải 8.760 giờ.")
        hourly_sum=[0.0]*24
        hourly_count=[0]*24
        for index,value in enumerate(ac):
            hour=index%24
            hourly_sum[hour]+=float(value or 0)
            hourly_count[hour]+=1
        hourly_average_kw=[round((hourly_sum[h]/hourly_count[h])/1000,6) for h in range(24)]
        station=data.get("station_info",{})
        return {"version":data.get("version"),"monthly":monthly,"annual":out.get("ac_annual"),
                "solrad_monthly":out.get("solrad_monthly"),"solrad_annual":out.get("solrad_annual"),
                "hourly_average_kw":hourly_average_kw,"hourly_sample_days":hourly_count,
                "hourly_source":"PVWatts V8 typical meteorological year · 8,760 hourly AC values",
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
        for typ,relation_target in relations:
            if relation_target in existing_targets: continue
            part=("xl/"+relation_target).replace("xl/../","")
            if part not in source.namelist(): continue
            ET.SubElement(rel_root,"{"+rel_ns+"}Relationship",{"Id":new_id(),"Type":typ,"Target":relation_target})
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

@app.post("/api/projects/{pid}/templates")
async def upload_project_templates(pid:str,request:Request,files:list[UploadFile]=File(...)):
    auth(request)
    rows=sb_request("GET",db("projects"),params={"id":f"eq.{pid}","select":"id,payload"}).json()
    if not rows or (rows[0].get("payload") or {}).get("_recordType")=="folder":
        raise HTTPException(404,"Không tìm thấy dự án.")
    xlsx=[f for f in files if Path(f.filename or "").suffix.lower()==".xlsx"]
    pptx=[f for f in files if Path(f.filename or "").suffix.lower()==".pptx"]
    if len(xlsx)!=1 or len(pptx)!=1:
        raise HTTPException(400,f"Thư mục phải có đúng 1 file .xlsx và 1 file .pptx. Hiện có {len(xlsx)} file Excel và {len(pptx)} file PowerPoint.")
    xdata=await xlsx[0].read(); pdata=await pptx[0].read()
    if not xdata or len(xdata)>10*1024*1024: raise HTTPException(413,"File PTTC rỗng hoặc vượt quá 10 MB.")
    if not pdata or len(pdata)>30*1024*1024: raise HTTPException(413,"File PowerPoint rỗng hoặc vượt quá 30 MB.")
    try:
        check=load_workbook(io.BytesIO(xdata),read_only=True,data_only=False)
        required={"THONG SO DAU VAO","present","draft ","16,66"}
        missing=required-set(check.sheetnames);check.close()
        if missing: raise ValueError("Thiếu sheet: "+", ".join(sorted(missing)))
    except Exception as exc: raise HTTPException(400,"File PTTC không hợp lệ: "+str(exc)[:250])
    try:
        from pptx import Presentation
        check=Presentation(io.BytesIO(pdata))
        if len(check.slides)<14: raise ValueError("Cần ít nhất 14 slide.")
    except Exception as exc: raise HTTPException(400,"File PowerPoint không hợp lệ: "+str(exc)[:250])
    base=f"projects/{pid}/templates"
    sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(base+"/financial.xlsx",safe="/"),data=xdata,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","x-upsert":"true"})
    sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(base+"/presentation.pptx",safe="/"),data=pdata,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.presentationml.presentation","x-upsert":"true"})
    return {"ok":True,"financial_name":Path(xlsx[0].filename or "PTTC.xlsx").name,"presentation_name":Path(pptx[0].filename or "Bao_gia.pptx").name,"slides":len(check.slides)}

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
        original=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(f"projects/{pid}/templates/financial.xlsx",safe="/")).content
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
        inverter_price=number("inverterPrice")
        if inverter_price is not None: inp["C5"]=inverter_price
        battery_price=number("batteryPrice")
        if battery_price is not None: inp["C6"]=battery_price
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
    pdf_filename=f"PTTC_{safe_name}_Trang_5-6.pdf"
    try:
        pdf_data=_xlsx_pages_to_pdf(generated,(5,6))
    except HTTPException:
        raise
    except Exception as exc:
        print(f"Financial PDF generation error: {type(exc).__name__}: {exc}",flush=True)
        raise HTTPException(500,"Không thể tạo PDF xem trước PTTC: "+str(exc)[:250])

    excel_id=str(uuid.uuid4()); excel_path=storage_path(pid,excel_id,filename)
    pdf_id=str(uuid.uuid4()); pdf_path=storage_path(pid,pdf_id,pdf_filename)
    uploaded_paths=[]
    try:
        sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(excel_path,safe="/"),data=generated,headers={"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","x-upsert":"true"})
        uploaded_paths.append(excel_path)
        sb_request("POST","/storage/v1/object/"+BUCKET+"/"+quote(pdf_path,safe="/"),data=pdf_data,headers={"Content-Type":"application/pdf","x-upsert":"true"})
        uploaded_paths.append(pdf_path)
        sb_request("POST",db("uploads"),json=[
            {"id":excel_id,"pid":pid,"name":filename,"path":excel_path},
            {"id":pdf_id,"pid":pid,"name":pdf_filename,"path":pdf_path}
        ],headers={"Prefer":"return=minimal"})
    except Exception:
        if uploaded_paths:
            sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":uploaded_paths})
        raise
    return {"id":excel_id,"name":filename,"excel_id":excel_id,"excel_name":filename,
            "pdf_id":pdf_id,"pdf_name":pdf_filename,
            "recalculation":"Đã tính lại công thức; PDF chỉ gồm 2 trang trong vùng đang hiển thị của sheet present."}


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


def _xlsx_pages_to_pdf(data, page_numbers=(5, 6)):
    """Export only the visible print area of the PTTC 'present' sheet.

    Hidden rows/columns in the template define which original pages are kept.
    Convert that visible rectangle into the actual print area so LibreOffice
    cannot paginate the hidden portions as extra pages.
    """
    import shutil, subprocess, tempfile
    from openpyxl.utils import get_column_letter, range_boundaries
    office=shutil.which("libreoffice") or shutil.which("soffice")
    if not office:
        raise HTTPException(503,"Máy chủ chưa có LibreOffice để xuất PDF PTTC.")

    try:
        wb=load_workbook(io.BytesIO(data))
        if "present" not in wb.sheetnames:
            raise HTTPException(422,"File PTTC không có sheet 'present'.")
        target=wb["present"]

        # Read the template's original print range, then trim away hidden
        # rows/columns. For the supplied template this resolves to J39:S146.
        if target.print_area:
            area=str(target.print_area)
            if "!" in area:
                area=area.rsplit("!",1)[1]
            area=area.replace("$","").replace("'","")
            min_col,min_row,max_col,max_row=range_boundaries(area)
        else:
            min_col,min_row,max_col,max_row=1,1,target.max_column,target.max_row

        visible_rows=[
            r for r in range(min_row,max_row+1)
            if not target.row_dimensions[r].hidden
        ]
        visible_cols=[
            c for c in range(min_col,max_col+1)
            if not target.column_dimensions[get_column_letter(c)].hidden
        ]
        if not visible_rows or not visible_cols:
            raise HTTPException(422,"Sheet present không có vùng in đang hiển thị.")

        first_row,last_row=min(visible_rows),max(visible_rows)
        first_col,last_col=min(visible_cols),max(visible_cols)
        target.print_area=(
            f"{get_column_letter(first_col)}{first_row}:"
            f"{get_column_letter(last_col)}{last_row}"
        )

        # Export only this sheet. Keep its page breaks, margins, orientation,
        # paper size and formatting from the user's template.
        for sheet in wb.worksheets:
            sheet.sheet_view.tabSelected=(sheet.title=="present")
            sheet.sheet_state="visible" if sheet.title=="present" else "hidden"
        wb.active=wb.sheetnames.index("present")

        # The retained visible area is one page wide and two pages tall in
        # the supplied template. Preserve manual row breaks (notably row 104).
        target.page_setup.fitToWidth=1
        target.page_setup.fitToHeight=2
        if target.sheet_properties.pageSetUpPr is None:
            from openpyxl.worksheet.properties import PageSetupProperties
            target.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
        else:
            target.sheet_properties.pageSetUpPr.fitToPage=True

        prepared=io.BytesIO()
        wb.save(prepared)
        wb.close()
        prepared_data=prepared.getvalue()
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(422,"Không đọc được bố cục in của sheet present: "+str(exc)[:200])

    with tempfile.TemporaryDirectory() as td:
        src=Path(td)/"PTTC.xlsx"
        outdir=Path(td)/"pdf"
        outdir.mkdir()
        src.write_bytes(prepared_data)
        proc=subprocess.run(
            [office,"-env:UserInstallation=file://"+str(Path(td)/"lo-pdf-profile"),
             "--headless","--convert-to","pdf","--outdir",str(outdir),str(src)],
            capture_output=True,text=True,timeout=180
        )
        pdf_path=outdir/"PTTC.pdf"
        if proc.returncode!=0 or not pdf_path.exists():
            raise HTTPException(500,"Không thể xuất PDF PTTC: "+(proc.stderr or proc.stdout or "LibreOffice PDF error")[:250])
        reader=PdfReader(str(pdf_path))
        if len(reader.pages)!=2:
            raise HTTPException(
                422,
                f"Vùng in đang hiển thị của sheet present tạo ra {len(reader.pages)} trang, không phải 2 trang. "
                "Vui lòng kiểm tra lại các dòng/cột đang ẩn và ngắt trang trong file mẫu PTTC."
            )
        writer=PdfWriter()
        writer.add_page(reader.pages[0])
        writer.add_page(reader.pages[1])
        output=io.BytesIO()
        writer.write(output)
        return output.getvalue()


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
        original=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(f"projects/{pid}/templates/presentation.pptx",safe="/")).content
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
        # Slide 3: update only self-consumption and 30-year savings from PTTC present!R133.
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
            accumulated=pttc_wb["present"]["R133"].value
        finally:
            pttc_wb.close()
        if accumulated is None:
            raise HTTPException(409,"Ô present!R133 chưa có kết quả. Hãy tạo lại báo cáo PTTC rồi xuất PowerPoint.")
        savings_text=f"{float(accumulated)/1000:.1f}".replace(".",",")
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
        text_shape(prs.slides[5],9,f"Biểu đồ sản lượng điện mặt trời dự kiến sản xuất của\nhệ thống {fmt(pv)} kWp\nTổng sản lượng 1 năm: {fmt(annual_pttc,0)} kWh")
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
        gsa_files=[]
        for image_no,shape_index in enumerate((11,12,13,14),start=1):
            found=[u for u in uploads if str(u.get("name",""))==f"global-solar-atlas-{image_no}.png"]
            if not found:
                raise HTTPException(409,"Cần chụp đủ 4 ảnh Global Solar Atlas trước khi xuất PowerPoint trang 6.")
            if shape_index<len(prs.slides[5].shapes):
                obj=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(found[-1]["path"],safe="/"))
                _ppt_replace_picture(prs.slides[5],prs.slides[5].shapes[shape_index],obj.content)
                gsa_files.append(image_no)
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
                  const tableText=table?(table.innerText||table.textContent||''):'';
                  const ghiMatch=tableText.match(/\\bGHI\\b[\\s\\S]{0,100}?(\\d{3,5}(?:[.,]\\d+)?)/i);
                  const ghi=ghiMatch?Number(ghiMatch[1].replace(',','.')):null;
                  return {metrics:!!table,maps:maps.length,sunpath:!!chart,ghi:Number.isFinite(ghi)?ghi:null};
                }""")
                required=["metrics","satellite","sunpath","pvout"]
                status=await page.evaluate("""() => Object.fromEntries(['metrics','satellite','sunpath','pvout'].map(k=>[k,!!document.querySelector('[data-gsa-shot="'+k+'"]')]))""")
                ghi_value=found.get("ghi")
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
        return {"images":saved,"source_url":url,"count":len(saved),"detected_panels":status,"ghi_kwh_m2_year":ghi_value}
    except HTTPException: raise
    except Exception as exc:
        print(f"Global Solar Atlas capture error: {type(exc).__name__}: {exc}",flush=True)
        raise HTTPException(502,f"Global Solar Atlas capture failed: {type(exc).__name__}: {str(exc)[:350]}")
    
LEGACY_FOLDER_ID="legacy-unassigned"

@app.get("/api/folders")
def folders(request:Request):
    auth(request)
    rows=sb_request("GET",db("projects"),params={"select":"id,name,payload,updated","order":"updated.desc"}).json()
    actual=[r for r in rows if (r.get("payload") or {}).get("_recordType")=="folder"]
    folder_ids={r.get("id") for r in actual}
    # Show older projects that predate folder management, or whose parent folder
    # no longer exists, in a virtual recovery folder. Do not rewrite/delete records
    # while rendering the dashboard.
    legacy=[r for r in rows if (r.get("payload") or {}).get("_recordType")!="folder"
            and (r.get("payload") or {}).get("folderId") not in folder_ids]
    result=[{"id":r["id"],"name":(r.get("payload") or {}).get("folderName") or r.get("name") or "Thư mục",
             "date":(r.get("payload") or {}).get("folderDate",""),"updated":r.get("updated")} for r in actual]
    if legacy:
        result.append({"id":LEGACY_FOLDER_ID,"name":"Dự án cũ (khôi phục)","date":"","updated":None})
    return result

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
    if fid==LEGACY_FOLDER_ID:
        valid_folder_ids={r.get("id") for r in rows if (r.get("payload") or {}).get("_recordType")=="folder"}
        selected=[r for r in rows if (r.get("payload") or {}).get("_recordType")!="folder"
                  and (r.get("payload") or {}).get("folderId") not in valid_folder_ids]
    else:
        selected=[r for r in rows if (r.get("payload") or {}).get("folderId")==fid
                  and (r.get("payload") or {}).get("_recordType")!="folder"]
    return [{"id":r["id"],"name":r.get("name"),"customer":r.get("customer"),"updated":r.get("updated")} for r in selected]

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
    if folder_id!=LEGACY_FOLDER_ID:
        frows=sb_request("GET",db("projects"),params={"id":f"eq.{folder_id}","select":"id,payload"}).json()
        if not frows or (frows[0].get("payload") or {}).get("_recordType")!="folder":
            raise HTTPException(404,"Không tìm thấy thư mục. Hãy tải lại danh sách thư mục.")
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

@app.delete("/api/projects/{pid}/files/{fid}")
def delete_project_file(pid:str,fid:str,request:Request):
    auth(request)
    rows=sb_request("GET",db("uploads"),params={"id":f"eq.{fid}","pid":f"eq.{pid}","select":"id,path,name"}).json()
    if not rows: raise HTTPException(404,"Không tìm thấy ảnh trong dự án này.")
    item=rows[0]
    sb_request("DELETE","/storage/v1/object/"+BUCKET,json={"prefixes":[item["path"]]})
    sb_request("DELETE",db("uploads"),params={"id":f"eq.{fid}","pid":f"eq.{pid}"},headers={"Prefer":"return=minimal"})
    return {"ok":True,"id":fid}

@app.get("/api/projects/{pid}/files")
def files(pid:str,request:Request):
    auth(request)
    r=sb_request("GET",db("uploads"),params={"pid":f"eq.{pid}","select":"id,name","order":"name.asc"})
    return r.json()

@app.get("/api/files/{fid}/preview")
def preview_file(fid:str,request:Request):
    auth(request)
    r=sb_request("GET",db("uploads"),params={"id":f"eq.{fid}","select":"name,path"})
    rows=r.json()
    if not rows: raise HTTPException(404,"Không tìm thấy file")
    item=rows[0]
    if not str(item["name"]).lower().endswith(".pdf"):
        raise HTTPException(415,"Chỉ hỗ trợ xem trước PDF.")
    obj=sb_request("GET","/storage/v1/object/"+BUCKET+"/"+quote(item["path"],safe="/"))
    return Response(content=obj.content,media_type="application/pdf",
                    headers={"Content-Disposition":f"inline; filename*=UTF-8''{quote(item['name'])}","Cache-Control":"no-store"})


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
