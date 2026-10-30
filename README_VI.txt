SOLAR QUOTE AUTOMATION - ONLINE MVP V1

CHUC NANG: dang nhap quan tri; tao/luu/mo/xoa du an tren may chu; nhap khach hang, cong trinh, cau hinh PV/inverter/pin, dien nang va gia; tinh kWp, dung luong pin, VAT; upload va tai lai PDF/Excel/CSV/anh/Word; xuat JSON.

DEPLOY RENDER:
1. Tao tai khoan GitHub va Render.
2. Tao repository GitHub moi, upload toan bo noi dung thu muc nay (giu nguyen app/main.py).
3. Render Dashboard > New > Blueprint, ket noi repository.
4. Render doc render.yaml; xac nhan tao service.
5. Khi duoc yeu cau, dat ADMIN_PASSWORD manh, rieng biet (toi thieu 16 ky tu).
6. Cho deploy thanh cong, mo URL onrender.com.
7. Ten dang nhap mac dinh admin; dung mat khau ADMIN_PASSWORD da cai.

CHI PHI / LUU TRU:
- render.yaml chon Starter web service va persistent disk 1GB; day la dich vu co phi. Kiem tra gia hien hanh truoc khi xac nhan.
- Persistent disk giu SQLite va file upload qua deploy. Khong xoa disk hoac doi DATA_DIR sau khi co du lieu.
- Sao luu du lieu dinh ky.
- MVP chi co mot tai khoan quan tri dung chung; chua co tai khoan rieng/phan quyen nhan vien.
- Chua co OCR hoa don, phu tai 30 phut, NPV/IRR, Excel/PPTX, ghep PDF.
- Tuyet doi khong dua mat khau/SESSION_SECRET len GitHub.
