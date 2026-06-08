# 📈 Multi-Bank Exchange Rate Crawler & Dashboard (USD/VND)

Hệ thống tự động hóa cào tỷ giá từ các ngân hàng lớn của Việt Nam (**Vietcombank, VietinBank, BIDV, Techcombank, ACB, SeaBank**) và tổng hợp vào Dashboard Excel (`TyGia_Banking.xlsx`) chuyên nghiệp với khả năng lưu trữ lịch sử, so sánh ngày chọn, định dạng điều kiện tự động (Conditional Formatting), vẽ biểu đồ trực quan và báo cáo thống kê tháng.

Tự động hóa hoàn toàn lịch trình chạy hằng ngày (18h00) qua Windows Task Scheduler.

---

## ✨ Tính Năng Nổi Bật

- **Crawl Dữ Liệu Đa Ngân Hàng (Pure Python - Không Dùng Playwright)**:
  - Tải **toàn bộ 20+ ngoại tệ** từ XML API của Vietcombank.
  - Tải tỷ giá **USD/VND** của **BIDV** (JSON API), **Techcombank** (JSON Integration API), **ACB** (REST API).
  - Tải tỷ giá **VietinBank & SeaBank** bằng request HTTPS trực tiếp sử dụng session CSRF token, không dùng Playwright giúp hệ thống chạy nhanh, tin cậy và tốn cực ít tài nguyên.
- **Dashboard Excel Trực Quan**:
  - **`TheoDoi_USD`**: Bảng điều khiển so sánh ngày nâng cao với Date Picker (B4) để xem bất kỳ ngày nào trong lịch sử, bộ chọn ngày so sánh (E4) hiển thị chênh lệch giá trị và % biến động kèm biểu đồ cột tự động cập nhật.
  - **`TheoDoi_Thang_USD`**: Trang báo cáo theo tháng, chọn Năm, Tháng, Ngân hàng để xem lịch sử tất cả các ngày trong tháng và tính tỷ giá bình quân cả tháng.
  - **`Data` & `Data_TheoDoi_USD`**: Sheet lưu trữ dữ liệu gốc dạng cộng dồn (append-only database).
- **Trực Quan Hóa Tối Ưu**:
  - Highlight tự động ngân hàng mua cao nhất (Xanh lá) và bán thấp nhất (Vàng) giúp người dùng tối ưu hóa giao dịch.
  - Định dạng điều kiện đổi màu chữ chênh lệch tăng (Đỏ) và giảm (Xanh lá).
- **Tính Năng Hệ Thống**:
  - **PDF Exporter**: Xuất báo cáo dạng PDF với thiết kế gọn gàng, chuyên nghiệp.
  - **Windows Task Scheduler Script**: Script PowerShell tự động cấu hình lịch chạy 18h00 mỗi ngày, cơ chế tự động thử lại (Retry) 3 lần nếu có lỗi kết nối.
  - **Tự Phục Hồi & Bảo Vệ Tài Nguyên**: Tự động tắt Excel trước khi ghi file, kiểm tra dung lượng ổ đĩa, tự động dọn dẹp file temp, và tự khôi phục dữ liệu từ bản backup gần nhất nếu file chính bị lỗi.

---

## 📂 Cấu Trúc Thư Mục Dự Án

```
vcb-exchange-rate/
├── docs/
│   └── ARCHITECTURE.md          # Chi tiết thiết kế hệ thống, công thức Excel
├── scripts/
│   ├── get_rates.py            # Script cào tỷ giá hàng ngày & cập nhật Excel Dashboard
│   ├── fetch_historical_all_banks.py # Script backfill lịch sử từ 01/01/2026 đến nay
│   ├── export_pdf.py           # Xuất báo cáo tỷ giá PDF chuyên nghiệp
│   ├── setup_schedule.ps1      # PowerShell Script cấu hình Task Scheduler chạy tự động
│   ├── test_excel.py           # Kịch bản kiểm thử tự động toàn bộ cấu trúc Excel
│   ├── package.py              # Script đóng gói dự án thành file ZIP sạch để release
│   └── fonts/
│       └── DejaVuSans.ttf      # Font hỗ trợ UTF-8 cho việc xuất PDF
├── README.md                   # Tài liệu hướng dẫn sử dụng (File này)
├── requirements.txt            # Danh sách thư viện Python cần thiết
└── LICENSE                     # Giấy phép MIT
```

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy Hệ Thống

### Điều Kiện Cần
- Hệ điều hành: Windows (để dùng Task Scheduler cấu hình tự động chạy).
- Python 3.8 trở lên.
- Microsoft Excel (để xem Dashboard trực quan).

### 1. Tải Dự Án & Cài Đặt Thư Viện
Tải mã nguồn về máy tính và cài đặt các thư viện cần thiết:

```bash
git clone https://github.com/your-username/vcb-exchange-rate.git
cd vcb-exchange-rate

# Tạo môi trường ảo (Khuyến nghị)
python -m venv .venv
.venv\Scripts\activate

# Cài đặt thư viện
pip install -r requirements.txt
```

### 2. Khởi Tạo Thư Mục Lưu Trữ
Dự án được cấu hình mặc định hoạt động tại thư mục cố định `D:\Tygia-Tudong` để đảm bảo lưu trữ an toàn, tránh phân mảnh và không ảnh hưởng ổ đĩa hệ điều hành:

```powershell
New-Item -ItemType Directory -Force -Path "D:\Tygia-Tudong"
```

### 3. Backfill Dữ Liệu Lịch Sử Từ 01/01/2026
Chạy script để tải toàn bộ tỷ giá lịch sử từ ngày 01/01/2026 đến nay và khởi tạo file Excel:

```bash
python scripts/fetch_historical_all_banks.py
```
*Sau khi hoàn tất, file `D:\Tygia-Tudong\TyGia_Banking.xlsx` sẽ chứa đầy đủ dữ liệu lịch sử và các trang Dashboard.*

### 4. Kiểm Tra Tính Toàn Vẹn Của Excel
Chạy test suite để chắc chắn cấu trúc các bảng biểu, dropdown, công thức và style hoạt động hoàn hảo:

```bash
python scripts/test_excel.py
```

### 5. Cài Đặt Task Scheduler Chạy Tự Động Hằng Ngày
Mở **PowerShell với quyền Administrator** và chạy script đăng ký Task Scheduler:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\scripts\setup_schedule.ps1
```
Tác vụ `TyGiaBanking_Daily` sẽ được thêm vào hệ thống để chạy tự động vào **18h00 mỗi ngày** với các thiết lập:
- Tự động chạy lại (Retry) 3 lần mỗi 10 phút nếu máy tính mất kết nối mạng.
- Tự động chạy bổ sung ngay khi mở máy nếu bỏ lỡ lịch hẹn trước đó.

---

## 📊 Mô Tả Báo Cáo Tháng `TheoDoi_Thang_USD`

Để xem dữ liệu thống kê cả tháng của một ngân hàng cụ thể:
1. Mở file `TyGia_Banking.xlsx`.
2. Truy cập sheet **`TheoDoi_Thang_USD`**.
3. Chọn **Năm** (ô `C4`), **Tháng** (ô `F4`), và **Ngân hàng** cần xem (ô `H4`) thông qua các Dropdown.
4. Bảng tính sẽ tự động đối chiếu dữ liệu lịch sử và hiển thị tỷ giá từng ngày trong tháng, đồng thời dòng số 38 sẽ tính toán **Tỷ giá bình quân cả tháng** của ngân hàng đó.

---

## 🛡️ Giấy Phép (License)

Dự án được phân phối dưới giấy phép MIT License.
