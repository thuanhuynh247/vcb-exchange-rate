# 📈 Multi-Bank Exchange Rate Crawler & Executive Dashboard (USD/VND)

Hệ thống tự động hóa cào tỷ giá từ **15 ngân hàng lớn nhất Việt Nam** (**Vietcombank, VietinBank, BIDV, Agribank, Techcombank, ACB, Sacombank, TPBank, VPBank, HDBank, Eximbank, OCB, SeaBank, VietABank, PVcomBank** - đầy đủ bộ tứ trụ **BIG 4** nhà nước và toàn bộ khối ngân hàng TMCP bán lẻ hàng đầu) và tổng hợp vào Dashboard Excel (`TyGia_Banking.xlsx`) chuyên nghiệp cùng **GitHub Pages Live Dashboard** trực quan, hiện đại.

🌐 **Trải nghiệm trực tiếp trên GitHub Pages**: [https://thuanhuynh247.github.io/vcb-exchange-rate/](https://thuanhuynh247.github.io/vcb-exchange-rate/)

Tự động hóa hoàn toàn lịch trình chạy hằng ngày (**23h00**) qua Windows Task Scheduler nhằm chốt chuẩn xác 100% tỷ giá đóng cửa trong ngày.

---

## ✨ Tính Năng Nổi Bật

- **Dữ Liệu Lịch Sử Toàn Diện 2026 (01/01/2026 - 09/09/2026)**:
  - Lưu trữ đầy đủ **252 ngày giao dịch liên tục** cho toàn bộ 15 ngân hàng ($252 \times 15 = 3,780$ bản ghi tỷ giá chuẩn mực).
  - Không có ngày khuyết thiếu, áp dụng chuẩn ngoại suy tài chính (forward/backward fill ngày nghỉ/lễ) đồng bộ với biến động thị trường.
- **Nâng Cấp Toàn Diện UI/UX Pro Max Live Dashboard (`index.html` & `docs/index.html`)**:
  - **Bento Grid & Glassmorphism**: Thiết kế giao diện hiện đại theo chuẩn Fintech Dashboard.
  - **Chế Độ Sáng / Tối (Dark & Light Mode)**: Chuyển đổi mượt mà với bộ lưu cấu hình `localStorage`.
  - **100% SVG Icons Chuẩn Mực**: Thay thế toàn bộ emoji icon bằng vector SVG sắc nét (Heroicons / Lucide).
  - **Bộ Quy Đổi Ngoại Tệ Trực Tiếp (Interactive Currency Converter)**: Nhập số tiền USD để tính toán số tiền VND nhận được hoặc phải thanh toán ngay tức thì tại 15 ngân hàng, tự động đề xuất ngân hàng có lợi nhất.
  - **Bảng So Sánh Matrix Pro Max**: Sắp xếp đa cột (Click tiêu đề cột để sort), huy hiệu Top 1/2/3 (Vàng, Bạc, Đồng), tìm kiếm ngân hàng trực tiếp theo thời gian thực, dòng `TRUNG BÌNH THỊ TRƯỜNG`.
  - **Bộ Biểu Đồ Kép (Dual Chart.js Suite)**:
    - Biểu đồ cột Flat UI 3 mức giá (Mua TM: Xanh ngọc, Mua CK: Xanh dương, Bán: Đỏ san hô).
    - Biểu đồ đường toàn cảnh 2026 (252 ngày) với bộ chọn thời gian linh hoạt (7N, 1T, 3T, 252N).
  - **Data Export Suite**: Nút xuất file CSV, nút tải file Excel `TyGia_Banking.xlsx`, nút tải API JSON `rates_history.json`, và in ấn / xuất PDF tối ưu layout.
- **Hỗ Trợ Phân Tích Chuyên Sâu Với Extension `vscode-smart-data-viewer`**:
  - Đã tích hợp sẵn gói extension `vnstock/vscode-smart-data-viewer` (v0.6.0) trong thư mục `tools/` và cài đặt tự động vào IDE.
  - Cho phép mở trực tiếp các file `TyGia_Banking.xlsx`, `rates_history.json`, và file CSV để chạy truy vấn DuckDB SQL, phân tích EDA và biểu đồ chuyên sâu ngay trong trình soạn thảo mã nguồn.
- **Crawl Dữ Liệu 15 Ngân Hàng (Pure Python - Không Dùng Playwright)**:
  - Tải **toàn bộ 20+ ngoại tệ** từ XML API của Vietcombank.
  - Tải tỷ giá **USD/VND** của **BIDV** (JSON API), **Techcombank** (JSON Integration API), **ACB** (REST API), **Agribank** (Cổng thông tin & WCM API), **VietABank** (Cổng niêm yết chính thức), **PVcomBank** (API JSON chính thức `Date=YYYY-MM-DD`).
  - Tải tỷ giá **VietinBank & SeaBank** bằng request HTTPS trực tiếp sử dụng Next.js Server Action; **VPBank, Sacombank, TPBank, Eximbank, HDBank, OCB** qua bộ trích xuất chuẩn hóa cao, hoàn toàn thuần Python không dùng Playwright giúp hệ thống chạy siêu nhanh (dưới 5 giây), tin cậy và tốn cực ít RAM.
- **Executive Financial Dashboard (Excel UI/UX Chuẩn Mực)**:
  - **4 Thẻ KPI Metric Cards Nổi Bật**: Tự động tính toán và hiển thị các mức giá tốt nhất thị trường.
  - **Bảng Ma Trận So Sánh 15 Ngân Hàng (`TheoDoi_USD`)**: Dropdown chọn ngày nhận diện toàn bộ 252 ngày, phân nhóm BIG 4 và TMCP, dòng `TRUNG BÌNH THỊ TRƯỜNG`, Conditional Formatting hiện đại.
  - **`TheoDoi_Thang_USD`**: Báo cáo theo tháng với bộ lọc dropdown hỗ trợ cả 15 ngân hàng.
- **Tính Năng Hệ Thống**:
  - **PDF Exporter**: Xuất báo cáo dạng PDF với thiết kế gọn gàng, chuyên nghiệp.
  - **Windows Task Scheduler Script**: Script PowerShell tự động cấu hình lịch chạy 23h00 mỗi ngày, cơ chế tự động thử lại (Retry) 3 lần nếu có lỗi kết nối.
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
Tác vụ `TyGiaBanking_Daily` sẽ được thêm vào hệ thống để chạy tự động vào **23h00 mỗi ngày** với các thiết lập:
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
