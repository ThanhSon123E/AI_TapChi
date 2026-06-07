# 📖 HƯỚNG DẪN CÀI ĐẶT VÀ CẤU HÌNH HỆ THỐNG (INSTALLATION & CONFIGURATION GUIDE)
> **Dự án:** AI Magazine Generator – Hệ thống tự động dàn trang và chuyển đổi tài liệu văn bản thành tạp chí số.
> **Sinh viên thực hiện:** Nguyen Thanh Son (NCTU)
> **Môi trường:** Python (Flask Framework) & MySQL (XAMPP Server)

---

## 🛠️ PHẦN I: TRIỂN KHAI CÁC BƯỚC CƠ BẢN (4 BƯỚC)

### Bước 1: Chuẩn bị Cơ sở dữ liệu (MySQL)
1. Khởi động phần mềm **XAMPP Control Panel**.
2. Bấm nút **Start** hai dịch vụ: `Apache` và `MySQL`.
3. Truy cập vào đường dẫn quản lý: `http://localhost/phpmyadmin/`.
4. Tạo một cơ sở dữ liệu mới đặt tên chính xác là: `ai_tapchi`.
5. Chọn database `ai_tapchi` ➔ Chọn tab **Import** ➔ Chọn file `ai_tapchi.sql` đi kèm trong thư mục dự án ➔ Kéo xuống dưới cùng bấm nút **Import**.

### Bước 2: Khởi tạo và Cài đặt Thư viện (Môi trường ảo)
Mở **Terminal / CMD** ngay tại thư mục gốc của dự án (`AI_Tapchi_GitHub`) và chạy lần lượt các lệnh sau:
```bash
# 1. Khởi tạo môi trường ảo Python độc lập
py -m venv venv

# 2. Kích hoạt môi trường ảo trên Windows
venv\Scripts\activate
# (Xuất hiện ký hiệu (venv) ở đầu dòng lệnh là thành công)

# 3. Cài đặt đồng loạt tất cả các thư viện phụ thuộc
pip install -r requirements.txt
Bước 3: Cấu hình Biến môi trường (.env)
Trong thư mục gốc của dự án, nhân bản file .env.example và đổi tên thành file: .env

Mở file .env lên, dán cấu hình mẫu bên dưới vào và tiến hành điền các mã khóa bảo mật theo hướng dẫn ở PHẦN II:

Ini, TOML
APP_NAME="AI_tapchi"
GOOGLE_CLIENT_ID=your_google_client_id_here
GOOGLE_CLIENT_SECRET=your_google_client_secret_here
SECRET_KEY=AITapChi

DB_HOST=localhost
DB_PORT=3306
DB_NAME=ai_tapchi
DB_USER=root
DB_PASSWORD=

# SePay Configuration
SEPAY_BANK_NAME=your_bank_name_here
SEPAY_ACCOUNT_NUMBER=your_bank_account_number_here
SEPAY_ACCOUNT_NAME=your_bank_account_name_here
Bước 4: Khởi chạy Ứng dụng
Chạy các lệnh cuối cùng này trên Terminal để bắt đầu vận hành hệ thống:

Bash
# 1. Khởi chạy Server Flask chính thức
py app.py
Đường dẫn truy cập: http://localhost:5000

Tài khoản Admin mặc định: admin@gmail.com / Mật khẩu: admin123

🔑 PHẦN II: HƯỚNG DẪN CHI TIẾT CÁCH LẤY API KEY
1. Hướng dẫn lấy Key đăng nhập Google (GOOGLE_CLIENT_ID & GOOGLE_CLIENT_SECRET)
Để kích hoạt tính năng đăng nhập nhanh bằng tài khoản Google, thực hiện các bước sau:

Truy cập vào trang quản lý: Google Cloud Console và đăng nhập bằng Gmail của bạn.

Bấm vào ô chọn dự án ở thanh trên cùng ➔ Chọn New Project (Dự án mới) ➔ Đặt tên dự án là AI-Magazine ➔ Bấm Create.

Tại menu bên trái, tìm đến mục APIs & Services ➔ Chọn OAuth consent screen (Màn hình chấp thuận OAuth):

Chọn User Type là External ➔ Bấm Create.

Điền thông tin bắt buộc: App name (AI Tạp Chí), User support email (Gmail của bạn) và Developer contact information (Gmail của bạn) ➔ Bấm Save and Continue qua hết các bước để hoàn tất.

Chuyển sang mục Credentials (Thông tin xác thực) ở menu bên trái:

Bấm nút + Create Credentials ở thanh trên cùng ➔ Chọn OAuth client ID.

Mục Application type (Loại ứng dụng): Chọn Web application.

Mục Name: Đặt tên gợi nhớ (Ví dụ: Flask Web App).

Cấu hình đường dẫn điều hướng (Quan trọng nhất):

Tại mục Authorized JavaScript origins, bấm Add URI và điền: http://localhost:5000

Tại mục Authorized redirect URIs (Đường dẫn nhận token xử lý đăng nhập), bấm Add URI và điền chính xác đường dẫn callback của mã nguồn: http://localhost:5000/login/google/callback

Bấm nút Create. Hệ thống sẽ hiển thị một cửa sổ chứa:

Your Client ID ➔ Copy dán vào dòng GOOGLE_CLIENT_ID= trong file .env.

Your Client Secret ➔ Copy dán vào dòng GOOGLE_CLIENT_SECRET= trong file .env.

💳 2. Hướng dẫn cấu hình cổng nạp tiền tự động SePay Webhook
Để hệ thống có thể tự động quét mã QR và cộng số dư tài khoản cho sinh viên khi họ nạp tiền, bạn cấu hình như sau:

A. Cấu hình thông tin ví trong file .env:
Điền trực tiếp thông tin ngân hàng của bạn (hoặc tài khoản ngân hàng tích hợp trên SePay) vào 3 dòng cuối của file .env:

SEPAY_BANK_NAME: Tên viết tắt của ngân hàng nhận tiền (Ví dụ: MBBank, Vietcombank, Techcombank...).

SEPAY_ACCOUNT_NUMBER: Số tài khoản ngân hàng chính xác nhận tiền của bạn.

SEPAY_ACCOUNT_NAME: Tên chủ tài khoản viết hoa không dấu (Ví dụ: NGUYEN THANH SON).

B. Đăng ký Webhook gửi dữ liệu về Web Đồ án:
Đăng nhập vào trang quản trị: SePay App.

Thêm ngân hàng của bạn vào hệ thống SePay để kết nối đọc lịch sử biến động số dư.

Vào mục Cấu hình Webhook (Tích hợp hệ thống) ➔ Bấm Tạo Webhook mới.

Cấu hình thông số Webhook:

URL Webhook: Điền đường dẫn API nhận dữ liệu biến động từ SePay của dự án. (Lưu ý: Nếu đang chạy dưới dạng Localhost trên máy tính, bạn cần sử dụng công cụ như Ngrok để tạo một link công khai proxy, ví dụ: https://xxxx-xx-xx.ngrok-free.app/webhook/sepay).

Phương thức (Method): Chọn POST.

Kiểu dữ liệu (Content-Type): Chọn json.

Bấm Lưu lại. Khi có bất kỳ ai quét mã QR hiển thị tại mục Billing trên trang web đồ án để nạp tiền, SePay sẽ tự động bắn một gói tin chứa số tiền nạp về hàm xử lý của hệ thống backend xử lý cộng tiền theo thời gian thực (Real-time).
