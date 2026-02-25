# Calendar Schedule App

Ứng dụng web hiển thị và quản lý lịch huấn luyện từ file Excel.

## Cài đặt trên Ubuntu

Dưới đây là các bước để cài đặt và chạy ứng dụng trên môi trường Ubuntu.

### 1. Cài đặt các phần mềm hệ thống cần thiết

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

### 2. Thiết lập Backend
Backend sử dụng Flask và chạy trên cổng `5001`.

1. Di chuyển vào thư mục backend:
```bash
cd backend
```
2. Tạo môi trường ảo (virtual environment) và kích hoạt nó:
```bash
python3 -m venv venv
source venv/bin/activate
```
3. Cài đặt các thư viện cần thiết:
```bash
pip install Flask flask-cors openpyxl werkzeug
```
4. Khởi động server backend:
```bash
python3 app.py
```
Backend sẽ chạy tại địa chỉ: `http://localhost:5001`

### 3. Thiết lập Frontend
Frontend chỉ bao gồm các file tĩnh (HTML, CSS, JS). Bạn cần một web server đơn giản để phục vụ file `index.html`.

1. Mở một terminal mới (giữ terminal backend vẫn đang chạy).
2. Di chuyển vào thư mục gốc của dự án (thư mục chứa `index.html`).
3. Khởi động HTTP server mặc định của Python (chạy trên cổng `8000`):
```bash
python3 -m http.server 8000
```
4. Mở trình duyệt và truy cập:
`http://localhost:8000` hoặc `http://127.0.0.1:8000`

---
**Lưu ý:**
- API backend được cấu hình sẵn để nhận request với CORS từ frontend.
- Cấu trúc thư mục chứa thư mục `backend/uploads/` sẽ được tự động tạo để lưu trữ các file Excel tải lên.
- Tên tuần (khi tải file lên) sẽ được dùng làm tên file Excel trên server.
