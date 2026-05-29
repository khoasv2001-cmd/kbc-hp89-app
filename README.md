# KBC-HP89 — Hệ thống quản lý liên kết KBC ↔ HP89

App Flask quản lý quy trình liên kết giữa **KBC** và **HP89**.

## ✨ Tính năng

- 📊 **Tổng Quan** — Dashboard tổng hợp tất cả mục
- 📦 **Đặt Hàng** — Workflow 5 bước:
  1. HP89 lưu nháp đơn hàng
  2. HP89 gửi → Chờ Lãnh đạo HP89 duyệt
  3. Lãnh đạo HP89 duyệt → đẩy sang KBC
  4. Nhân viên KBC tick "Nhận đơn"
  5. Nhân viên KBC tick "Đã giao hàng" → hoàn thành
- 📄 **Hợp Đồng** — Lưu hợp đồng, thỏa thuận, biên bản, hóa đơn giữa KBC & HP89
- ⚖️ **Pháp Lý** — Cây giấy tờ pháp lý nhiều cấp, phê duyệt bởi lãnh đạo
- 📢 **Truyền Thông** — HP89 đăng nội dung báo chí/truyền thông → KBC duyệt thống nhất
- 📥 **Tải Báo Cáo** — Xuất Excel cho từng danh mục
- 👥 **Người Dùng** — Phân biệt nhân sự **KBC** ✅ và **HP89** 🔵, phân quyền chi tiết

## 🚀 Chạy trên Windows

Mở thư mục `kbc_hp89_app`, **bấm đôi vào `run.bat`**.
- Lần đầu sẽ tạo venv + cài thư viện (3-5 phút).
- App tự mở trình duyệt: <http://127.0.0.1:5000>
- Tài khoản admin: **admin / admin123** → đổi mật khẩu ngay sau khi đăng nhập.

## 📱 Cài đặt như app trên điện thoại

App đã hỗ trợ **PWA**:
- Mở Chrome/Edge trên điện thoại, vào địa chỉ app
- Menu trình duyệt → "Cài đặt ứng dụng" / "Add to home screen"

## ☁️ Deploy lên Render (miễn phí)

1. Tạo repo GitHub, push toàn bộ thư mục này lên.
2. Vào <https://render.com> → "New Web Service" → kết nối GitHub repo.
3. Build command: `pip install -r requirements.txt`
4. Start command: tự lấy từ `Procfile`
5. Tạo Persistent Disk, mount vào `/var/data`, đặt biến môi trường `DATA_DIR=/var/data` để dữ liệu không mất khi redeploy.

## 🔔 Bật thông báo nổi (Web Push)

Tạo cặp khóa VAPID:
```bash
pip install py-vapid
vapid --gen
```
Đặt vào biến môi trường (Render Dashboard):
- `VAPID_PUBLIC_KEY`
- `VAPID_PRIVATE_KEY`
- `VAPID_CLAIM_EMAIL` = `mailto:admin@kbc-hp89.vn`

## 📧 Gửi email (tùy chọn)

Đặt biến môi trường:
- `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`

## 📂 Cấu trúc

```
kbc_hp89_app/
├── app.py                  # App Flask chính
├── requirements.txt        # Thư viện
├── Procfile                # Deploy Render/Heroku
├── runtime.txt             # Python version
├── run.bat                 # Chạy local Windows
├── data.db                 # SQLite (tự sinh khi chạy)
├── static/
│   ├── kbc-logo.png        # Logo KBC
│   ├── style.css
│   ├── manifest.json       # PWA
│   └── sw.js               # Service worker
├── templates/              # Tất cả file giao diện HTML
└── uploads/                # File user upload
```

## ⚙️ Quyền của mỗi loại nhân sự

### HP89
- Nhân viên: tạo đơn hàng (nháp), soạn nội dung truyền thông
- Lãnh đạo / Phó GĐ: duyệt đơn hàng phía HP89, duyệt giấy tờ pháp lý

### KBC
- Nhân viên: nhận đơn hàng (đã được HP89 duyệt), tick đã giao hàng
- Lãnh đạo: duyệt nội dung truyền thông HP89 gửi, ký hợp đồng

### Admin (KBC)
- Toàn quyền hệ thống, quản lý người dùng, backup DB.
