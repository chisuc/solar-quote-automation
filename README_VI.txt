SOLAR QUOTE AUTOMATION - BẢN DÙNG RENDER FREE + SUPABASE FREE

1) Cập nhật toàn bộ repository GitHub bằng các file trong gói này:
   - app/main.py (giữ thư mục app)
   - requirements.txt
   - render.yaml
   - supabase_setup.sql (thêm file ở thư mục gốc)
2) Trong Supabase > SQL Editor > New query: mở supabase_setup.sql, dán toàn bộ SQL và nhấn Run.
3) Kiểm tra Storage bucket tên customer-files tồn tại và Public bucket đang OFF.
4) Trên Render, tạo Web Service từ repository, chọn Free. Thêm Environment Variables:
   ADMIN_USER=admin
   ADMIN_PASSWORD=<tự đặt mật khẩu đăng nhập mạnh>
   SESSION_SECRET=<chuỗi ngẫu nhiên dài>
   COOKIE_SECURE=true
   SUPABASE_URL=https://<project-ref>.supabase.co
   SUPABASE_SECRET_KEY=<secret key của Supabase>
   SUPABASE_BUCKET=customer-files
   Không đưa secret key vào GitHub, không chia sẻ secret key qua chat.
5) Deploy. Dùng URL Render để đăng nhập.

LƯU Ý: Render Free có thể sleep khi không hoạt động và cold start. Supabase Free có giới hạn dung lượng/traffic và chính sách tạm dừng project khi không hoạt động. Nên xuất dữ liệu/backup định kỳ.

Tài khoản đăng nhập website: ADMIN_USER / ADMIN_PASSWORD do bạn tự cấu hình trong Render.
