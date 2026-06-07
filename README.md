# 📖 HỆ THỐNG AI MAGAZINE GENERATOR (AI TẠP CHÍ)
> **Đề tài:** Nền tảng tự động dàn trang và chuyển đổi tài liệu văn bản thành tạp chí số sử dụng Trí tuệ nhân tạo (AI).
> **Sinh viên thực hiện:** Nguyen Thanh Son (NCTU)
> **Môi trường phát triển:** Python (Flask Framework) & MySQL (XAMPP Server)

---

## 🛠️ PHẦN I: TỔNG QUAN DỰ ÁN (PROJECT OVERVIEW)

### 1. Công nghệ sử dụng (Technology Stack)
* **Backend:** Python (`Flask`, `SQLAlchemy`). Tích hợp thư viện chuyên dụng `ReportLab` và `python-docx` để tự động bóc tách cấu trúc văn bản thô và dàn trang PDF ngầm (Background Job).
* **Frontend:** HTML5, CSS3 (Vanilla) và JavaScript. Tích hợp thư viện đồ họa nâng cao `page-flip.js` và `PDF.js` để mô phỏng giao diện đọc sách, lật trang 3D (Flipbook) trực tuyến sinh động.
* **Cơ sở dữ liệu:** Hỗ trợ đồng bộ hoàn toàn sang hệ quản trị cơ sở dữ liệu **MySQL (XAMPP)** cho môi trường triển khai thực tế.
* **Cổng thanh toán:** Tích hợp hệ thống tự động quét mã QR và kết nối qua `SePay Webhook` để nhận dữ liệu biến động số dư, tự động xử lý cộng tiền vào tài khoản người dùng theo thời gian thực (Real-time).
* **Trí tuệ nhân tạo (AI):** Hỗ trợ linh hoạt đa nền tảng LLM Provider (`OpenRouter`, `OpenAI`, `Gemini`, `DeepSeek`) qua kết nối API để hỗ trợ người dùng dịch thuật nội dung hoặc tối ưu hóa bố cục tạp chí tự động.

### 2. Quy trình thiết kế & Tính năng dành cho Người dùng
Hệ thống cung cấp một quy trình khép kín giúp người dùng phổ thông biến file văn bản thô thành tạp chí cao cấp:
* **Đăng nhập / Đăng ký:** Hỗ trợ tạo tài khoản nội bộ bảo mật hoặc đăng nhập nhanh chỉ với 1 click chuột qua hệ thống định danh `Google OAuth2`.
* **Nạp tiền tự động (Billing):** Người dùng thực hiện quét mã QR hiển thị trên hệ thống để nạp tiền thông qua cổng SePay. Mỗi lượt tạo tạp chí sẽ tiêu tốn một khoản phí cố định (mặc định cấu hình: `10,000 VND`).
* **Quy trình dàn trang tự động 5 bước:**
  1. **Bước 1 (Chọn Template):** Lựa chọn phong cách thiết kế tạp chí định sẵn: *VOGUE* (sang trọng), *MINIMAL* (tối giản), *ELEGANT* (thanh lịch), *MODERN* (hiện đại), hoặc *LUXURY* (cao cấp).
  2. **Bước 2 (Cấu hình bìa):** Tải lên tệp tin ảnh bìa trước (Front Cover), ảnh bìa sau (Back Cover) và nhập tiêu đề chính/phụ của ấn phẩm.
  3. **Bước 3 (Tải lên tài liệu):** Upload file `.docx` được định dạng chuẩn (hỗ trợ bóc tách chia cột, tạo chữ cái phóng to đầu dòng *Drop Cap*, đoạn trích dẫn nổi bật *Pullquote*, chèn hình ảnh minh họa và chú thích ảnh).
  4. **Bước 4 (Xử lý ngầm):** Hệ thống kích hoạt Background Job tự động đọc file, trích xuất hình ảnh, áp dụng thuật toán dàn trang bẻ đôi thành bố cục 2 cột, lồng ghép trang bìa và kết xuất thành tệp PDF hoàn chỉnh.
  5. **Bước 5 (Trải nghiệm & Tải về):** Đọc tạp chí trực tiếp với hiệu ứng lật trang 3D lướt mượt mà trên trình duyệt hoặc bấm nút tải tệp PDF chất lượng cao về máy tính.

### 3. Phân hệ Quản trị (Admin Dashboard)
Khi đăng nhập bằng tài khoản quản trị (Tài khoản mặc định: `admin@gmail.com` / Mật khẩu: `123456`), Admin có toàn quyền:
* **Quản lý người dùng:** Theo dõi toàn bộ danh sách thành viên hệ thống, thực hiện khóa/mở khóa tài khoản, hoặc chủ động cộng/trừ số dư ví thủ công cho người dùng.
* **Quản lý giao dịch:** Kiểm tra lịch sử tạo tạp chí, tiến trình bóc tách tệp và quản lý trạng thái các hóa đơn nạp tiền tự động.
* **Cấu hình hệ thống:** Thay đổi linh hoạt cấu hình kết nối AI (LLM Provider/Model), cập nhật danh sách API Keys, thiết lập giá tiền cho mỗi lượt tạo tạp chí và tùy chỉnh định dạng các gói nạp tiền bằng cấu hình dữ liệu `JSON`.

---

## ⚙️ PHẦN II: QUY TRÌNH TRIỂN KHAI VÀ CÀI ĐẶT (4 BƯỚC)

### 1. Chuẩn bị Cơ sở dữ liệu (MySQL)
1. Khởi động phần mềm **XAMPP Control Panel** trên máy tính.
2. Bấm nút **Start** hai dịch vụ: `Apache` và `MySQL` (Đảm bảo các ô chuyển sang màu xanh lá cây).
3. Mở trình duyệt web và truy cập vào đường dẫn: `http://localhost/phpmyadmin/`.
4. Tạo một cơ sở dữ liệu mới đặt tên chính xác là: `ai_tapchi`.
5. Chọn database `ai_tapchi` vừa tạo ➔ Chọn tab **Import** ➔ Bấm nút chọn file `ai_tapchi.sql` đi kèm trong bộ mã nguồn ➔ Kéo xuống dưới cùng bấm nút **Import** để hoàn tất nạp cấu trúc bảng biểu.


### 2. Khởi tạo và Cài đặt Thư viện (Môi trường ảo)
Mở cửa sổ **Terminal / CMD** ngay tại thư mục gốc của dự án và chạy lần lượt các câu lệnh sau:

```bash
# 1. Khởi tạo môi trường ảo Python độc lập (tạo thư mục venv sạch)
py -m venv venv

# 2. Kích hoạt môi trường ảo trên hệ điều hành Windows
venv\Scripts\activate
# (Dấu hiệu thành công: Xuất hiện chữ (venv) màu xanh ở đầu dòng lệnh Terminal)

# 3. Tiến hành cài đặt đồng loạt tất cả các thư viện phụ thuộc của dự án
pip install -r requirements.txt
```
3. Cấu hình Biến môi trường (.env)
Trong thư mục gốc của dự án, tiến hành nhân bản file .env.example và đổi tên thành file: .env

Mở file .env vừa tạo lên bằng phần mềm soạn thảo và điền đầy đủ cấu hình thông số mẫu như sau:
```bash
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
```
4. Khởi chạy Ứng dụng
Thực hiện chạy 2 câu lệnh cuối cùng này trên cửa sổ Terminal để kích hoạt hệ thống:
```bash
# 1. Khởi tạo tài khoản Quản trị (Admin) mặc định cho hệ thống
py create_admin.py

# 2. Khởi chạy Server Flask chính thức để mở máy chủ Web Local
py app.py
```
Đường dẫn truy cập: Mở trình duyệt web và truy cập địa chỉ địa phương: http://localhost:5000

Tài khoản Admin đăng nhập: admin@gmail.com / Mật khẩu: 123456

🔑 PHẦN III: HƯỚNG DẪN CHI TIẾT CÁCH LẤY API KEY CẤU HÌNH
1. Cách lấy mã đăng nhập Google OAuth2 (GOOGLE_CLIENT_ID & GOOGLE_CLIENT_SECRET)
Truy cập vào cổng quản lý: Google Cloud Console và đăng nhập bằng tài khoản Gmail cá nhân.

Bấm vào ô chọn dự án ở thanh menu trên cùng ➔ Chọn New Project ➔ Đặt tên dự án là AI-Magazine ➔ Bấm Create.

Tại danh sách menu bên trái, tìm chọn mục APIs & Services ➔ Chọn tiếp mục OAuth consent screen:

Chọn loại User Type là External ➔ Bấm Create.

Điền đầy đủ các thông tin bắt buộc gồm: App name (AI Tạp Chí), User support email (Gmail của bạn) và Developer contact information (Gmail của bạn) ➔ Bấm Save and Continue qua hết các bước sau để xác nhận hoàn tất.

Di chuyển sang mục Credentials (Thông tin xác thực) ở menu bên trái:

Bấm chọn nút + Create Credentials ở thanh công cụ phía trên ➔ Chọn dòng OAuth client ID.

Mục Application type (Loại ứng dụng): Chọn loại Web application.

Cấu hình đường dẫn điều hướng bắt buộc để chạy dưới Local (Cực kỳ quan trọng):

Tại mục Authorized JavaScript origins, bấm Add URI và điền chính xác: `http://localhost:5000`

Tại mục Authorized redirect URIs, bấm Add URI và điền chính xác đường dẫn hàm xử lý callback của mã nguồn: `http://localhost:5000/login/google/callback`

Bấm nút Create. Hệ thống sẽ xuất hiện một bảng thông báo hiển thị 2 chuỗi mã:

Your Client ID ➔ Copy chuỗi này dán đè vào mục GOOGLE_CLIENT_ID= trong file .env.

Your Client Secret ➔ Copy chuỗi này dán đè vào mục GOOGLE_CLIENT_SECRET= trong file .env.

2. Cách cấu hình cổng nạp tiền tự động SePay Webhook
Bước A: Điền thông tin ví cá nhân vào file .env
Cập nhật trực tiếp thông tin ngân hàng thụ hưởng nhận tiền của bạn vào 3 dòng cuối cùng trong file cấu hình .env:

SEPAY_BANK_NAME: Ghi tên viết tắt chuẩn của ngân hàng nhận tiền (Ví dụ: MBBank, Vietcombank, Techcombank...).

SEPAY_ACCOUNT_NUMBER: Ghi số tài khoản ngân hàng chính xác dùng để nhận tiền nạp.

SEPAY_ACCOUNT_NAME: Ghi họ và tên chủ tài khoản viết in hoa không dấu (Ví dụ: NGUYEN THANH SON).
`ngrok http --domain=versie-prodigious-imani.ngrok-free.dev 5000`

Bước B: Cấu hình Webhook kết nối dữ liệu
Đăng nhập vào trang quản trị: SePay App.

Thêm số tài khoản ngân hàng của bạn vào hệ thống SePay để cấp quyền đọc lịch sử biến động số dư tự động qua ứng dụng Mobile Banking.

Tìm đến mục Cấu hình Webhook (Tích hợp hệ thống) ➔ Bấm nút Tạo Webhook mới.

Cấu hình chính xác thông số truyền tải gói tin như sau:

URL Webhook: Điền link nhận dữ liệu từ SePay của dự án web (Nếu chạy ở Localhost, cần dùng công cụ như Ngrok để public link, ví dụ: `https://xxxx.ngrok-free.app/webhook/sepay).`

Phương thức (Method): Chọn POST.

Kiểu dữ liệu dữ liệu (Content-Type): Chọn định dạng cấu trúc dữ liệu json.

Bấm nút Lưu lại. Khi có bất kỳ tài khoản người dùng nào trên web thực hiện quét mã QR nạp tiền, SePay sẽ lập tức gửi dữ liệu biến động về máy chủ Backend xử lý phân tích cú pháp và tự động cộng tiền số dư tức thì theo thời gian thực.
