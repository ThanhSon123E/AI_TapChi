# AI_TapChi
# 📖 HƯỚNG DẪN CÀI ĐẶT VÀ KHỞI CHẠY HỆ THỐNG (INSTALLATION GUIDE)
> **Dự án:** AI Magazine Generator – Hệ thống tự động dàn trang và chuyển đổi tài liệu văn bản thành tạp chí số.
> **Sinh viên thực hiện:** Nguyen Thanh Son (NCTU)
> **Môi trường:** Python (Flask Framework) & MySQL (XAMPP Server)

---

## 🛠️ Bước 1: Chuẩn bị Cơ sở dữ liệu (MySQL)

1. Khởi động phần mềm **XAMPP Control Panel** trên máy tính.
2. Bấm nút **Start** hai dịch vụ: `Apache` và `MySQL` (Đợi các ô chuyển sang màu xanh lá cây).
3. Mở trình duyệt web và truy cập vào đường dẫn quản lý: `http://localhost/phpmyadmin/`.
4. Tạo một cơ sở dữ liệu mới (Database) đặt tên chính xác là: `ai_tapchi`.
5. Chọn database `ai_tapchi` vừa tạo ➔ Chọn tab **Import** ➔ Bấm *Choose File* và chọn file `ai_tapchi.sql` đi kèm trong thư mục dự án ➔ Kéo xuống dưới cùng bấm nút **Import** để hoàn tất nạp bảng biểu.

---

## 🚀 Bước 2: Khởi tạo và Cài đặt Thư viện (Môi trường ảo)

Mở **Terminal / CMD** ngay tại thư mục gốc của dự án (`AI_Tapchi_GitHub`) và tiến hành chạy lần lượt bộ lệnh sau:

```bash
# 1. Khởi tạo môi trường ảo Python độc lập (tạo thư mục venv)
py -m venv venv

# 2. Kích hoạt môi trường ảo trên hệ điều hành Windows
venv\Scripts\activate
# (Dấu hiệu thành công: Xuất hiện ký hiệu (venv) ở đầu dòng lệnh của Terminal)

# 3. Tiến hành cài đặt đồng loạt tất cả các thư viện phụ thuộc (Flask, ReportLab, python-docx,...)
pip install -r requirements.txt
