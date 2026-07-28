---
name: get-vcb-exchange-rate
description: |
  Tự động lấy tỷ giá ngoại tệ (tất cả 20+ ngoại tệ) từ trang web Vietcombank 
  và lấy tỷ giá USD của nhiều ngân hàng (VietinBank, BIDV, Techcombank, ACB, SeaBank) 
  để tổng hợp so sánh trực quan trong file Excel tại D:\Tygia-Tudong.
  Được cài đặt chạy tự động bằng Windows Task Scheduler.
---

# Goal
Tự động lấy tỷ giá **tất cả ngoại tệ** từ Vietcombank và tỷ giá **USD** từ nhiều ngân hàng lớn (VietinBank, BIDV, Techcombank, ACB, SeaBank). Lưu trữ cộng dồn và tạo Dashboard phân tích, so sánh trực quan. Toàn bộ tiến trình chạy tự động mỗi ngày thông qua Windows Task Scheduler.

# Instructions

## Script chính: `scripts/get_rates.py`
- Chạy tự động hằng ngày qua Task Scheduler (18h00) hoặc chạy thủ công.
- Lấy **tất cả 20+ ngoại tệ** từ XML API của VCB.
- Ghi dữ liệu cộng dồn vào sheet **Data** (có AutoFilter, dễ tìm kiếm).
- **Thu thập tỷ giá USD đa ngân hàng**:
  - **Vietcombank**: XML API.
  - **BIDV**: JSON API (`ExchangeDetailServlet`).
  - **Techcombank**: Integration JSON API.
  - **ACB**: REST API (`effectiveDateTime` filter).
  - **VietinBank & SeaBank**: Dùng Playwright/Chromium để crawl dữ liệu Real-time an toàn, bỏ qua các vấn đề chặn CORS/Cloudflare.
- Ghi dữ liệu USD cộng dồn ẩn vào sheet **Data_TheoDoi_USD**.
- Tạo sheet **TheoDoi_USD** nâng cao:
  - **Date Picker** (B4): dropdown chọn ngày xem tỷ giá.
  - **So sánh ngày** (E4): dropdown chọn ngày so sánh (hiển thị chênh lệch giá trị và % thay đổi).
  - **Tìm giá trị tốt nhất (Conditional Formatting)**: Tự động highlight ngân hàng có giá mua cao nhất (Xanh lá) và ngân hàng có giá bán thấp nhất (Vàng) giúp tối ưu giao dịch.
  - **Biểu đồ so sánh**: Cột biểu đồ BarChart so sánh trực quan giá Mua/Bán giữa 6 ngân hàng.
- **Tính năng tự phục hồi & Dọn dẹp**:
  - Tự động đóng Excel (`taskkill`) nếu file đang bị mở (tránh khóa file khi ghi).
  - Tự động kiểm tra dung lượng trống ổ D, nếu `< 500MB` tiến hành dọn dẹp thư mục Temp và các bản backup cũ để giải phóng không gian.
  - Tự khôi phục từ bản backup gần nhất nếu file chính bị lỗi hoặc 0 bytes.

## Script lịch sử: `scripts/fetch_historical.py`
- Chạy thủ công khi cần lấy dữ liệu lịch sử từ 2025-01-01.
- Sử dụng JSON API của VCB, hỗ trợ đa luồng (10 workers).
- Tự phát hiện dữ liệu cũ và chỉ lấy ngày còn thiếu.

## Export PDF: `scripts/export_pdf.py`
- Xuất báo cáo tỷ giá dạng PDF với branding VCB.
- `python export_pdf.py` → ngày mới nhất.
- `python export_pdf.py --date 10/03/2026` → ngày cụ thể.
- `python export_pdf.py --compare 01/03/2026` → so sánh với ngày khác.
- Output: `TyGia_Banking_YYYYMMDD.pdf` trong Documents.

## Script kiểm tra: `scripts/test_excel.py`
- Kiểm tra cấu trúc file Excel, format, Dashboard, Date Picker, so sánh 2 ngày, so sánh ngân hàng, biểu đồ, log file.

## Cài đặt Task Scheduler: `scripts/setup_schedule.ps1`
- Chạy bằng PowerShell (Admin) để đăng ký Task Scheduler.
- Định cấu hình chạy mỗi ngày lúc 18h00.
- Retry 3 lần mỗi 10 phút, timeout 15 phút, StartWhenAvailable.

# Examples
## File Excel Output (`D:\Tygia-Tudong\TyGia_Banking.xlsx`)

### Sheet "TheoDoi_USD"
- **Date Picker** ô B4: dropdown danh sách ngày lịch sử.
- **So sánh** ô E4: chọn ngày so sánh → cột G (Bán SS), H (Chênh lệch), I (% SS).
- **Conditional formatting**:
  - Ô có giá Mua cao nhất được tô màu xanh lá nhạt (#E8F5E9) với chữ xanh đậm (#2E7D32).
  - Ô có giá Bán thấp nhất được tô màu vàng (#FFEB3B) với chữ xanh ô-liu (#827717).
  - Cột Chênh lệch tự động đổi màu Đỏ khi tăng giá bán, Xanh lá khi giảm giá bán.
- **Biểu đồ**: Cột BarChart tự động cập nhật theo ngày được chọn ở Date Picker.

### Sheet "Data"
| Ngày Cập Nhật | Mã Ngoại Tệ | Tên Ngoại Tệ | Mua Tiền Mặt | Mua Chuyển Khoản | Bán     |
|---------------|-------------|---------------|--------------|------------------|---------|
| 2026-06-08    | USD         | US DOLLAR     | 26,097.00    | 26,127.00        | 26,407  |

# Constraints
- Đảm bảo máy tính bật và có kết nối mạng lúc 18h00.
- Đường dẫn cố định dự án: `D:\Tygia-Tudong`
- Excel File: `D:\Tygia-Tudong\TyGia_Banking.xlsx`
- Log File: `D:\Tygia-Tudong\vcb_rates.log`
- Backup folder: `D:\Tygia-Tudong\Backup`
- `export_pdf.py` cần thư viện `fpdf2` (`pip install fpdf2`). Font DejaVu tự tải lần đầu.
