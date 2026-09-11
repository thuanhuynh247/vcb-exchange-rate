---
name: get-vcb-exchange-rate
description: |
  Tự động thu thập tỷ giá ngoại tệ từ Vietcombank (tất cả 20+ loại ngoại tệ) 
  và tỷ giá USD/VND của 15 ngân hàng lớn nhất Việt Nam (Big 4 + TMCP hàng đầu). 
  Tổng hợp tự động vào Excel Dashboard (D:\Tygia-Tudong\TyGia_Banking.xlsx) 
  và triển khai Live Dashboard GitHub Pages theo chuẩn UI/UX Pro Max 2.1.
  Tự động hóa hoàn toàn lịch chạy 23h00 hàng ngày qua Windows Task Scheduler.
---

# Mục Tiêu Hệ Thống (Goal)
Thu thập, chuẩn hóa và phân tích tỷ giá **USD/VND** của **15 ngân hàng lớn nhất Việt Nam** (**Vietcombank, VietinBank, BIDV, Agribank, Techcombank, ACB, Sacombank, TPBank, VPBank, HDBank, Eximbank, OCB, SeaBank, VietABank, PVcomBank**) và **toàn bộ 20+ ngoại tệ** của Vietcombank. Lưu trữ lịch sử liên tục từ 01/01/2026 đến nay, cung cấp bảng phân tích tài chính Excel chuyên nghiệp và cổng thông tin trực quan trên GitHub Pages.

---

# Hướng Dẫn & Quy Trình Thực Thi (Instructions)

## 1. Script Cào Tỷ Giá Hằng Ngày: `scripts/get_rates.py`
- Chạy tự động vào **23h00 mỗi ngày** qua Windows Task Scheduler để chốt chính xác giá đóng cửa phiên.
- **Thu thập toàn diện 15 ngân hàng (Pure Python, không dùng Playwright)**:
  - **Vietcombank**: XML API (`Usercontrols/TVPortal.TyGia/pXML.aspx?b=68`).
  - **BIDV**: JSON API qua POST request (`ExchangeDetailServlet`).
  - **Techcombank**: Integration JSON API qua token CSRF.
  - **ACB**: REST API (`api/front/v1/currency?currency=VND`).
  - **Agribank**: Trích xuất bảng tỷ giá niêm yết cổng thông tin & API WCM.
  - **VietinBank & SeaBank**: Request HTTPS trực tiếp sử dụng Next.js Server Action (`next-action`).
  - **PVcomBank**: Cổng API niêm yết chính thức (`exchange-rate-by-date?Date=YYYY-MM-DD`).
  - **VietABank, VPBank, Sacombank, TPBank, Eximbank, HDBank, OCB**: Bộ trích xuất chuẩn hóa cao, tin cậy và tốn cực ít tài nguyên.
- **Ghi dữ liệu cộng dồn vào Excel (`TyGia_Banking.xlsx`)**:
  - Sheet `Data`: Toàn bộ 20+ ngoại tệ VCB.
  - Sheet `Data_TheoDoi_USD`: Bảng tỷ giá USD cộng dồn của 15 ngân hàng.
  - Sheet `TheoDoi_USD`: Executive Dashboard tương tác với Date Picker, Conditional Formatting, thẻ KPI và Biểu đồ.
- **Cơ chế tự phục hồi & An toàn**:
  - Đóng Excel tự động (`taskkill`) trước khi ghi file nhằm tránh khóa file.
  - Kiểm tra dung lượng ổ đĩa D, dọn dẹp thư mục Temp và bản backup cũ nếu dung lượng `< 500MB`.
  - Tự động khôi phục từ bản backup gần nhất nếu file chính bị lỗi.

## 2. GitHub Pages Live Dashboard: `index.html` & `docs/index.html`
- **Chuẩn Thiết Kế UI/UX Pro Max 2.1**:
  - **Hệ số chuẩn quốc tế**: Sử dụng dấu phẩy phân tách hàng nghìn `,000,000 VND` (`26,099,000 VND`, `25,720 VND`).
  - **Bộ công cụ sao chép 1-click**: Hỗ trợ nhấp để copy từng ô tỷ giá, copy theo dòng ngân hàng, copy thẻ KPI, copy kết quả quy đổi, xuất bảng Excel (TSV) và tóm tắt văn bản gửi Zalo/Telegram.
  - **Bento Grid & OLED Dark/Light Mode**: Tự động ghi nhớ cấu hình giao diện vào `localStorage`.
  - **Bank Identity Badges**: Huy hiệu nhận diện 15 ngân hàng với màu thương hiệu chính xác.
  - **Arbitrage Opportunity Card**: Tự động phát hiện cặp ngân hàng có biên độ chênh lệch giá tốt nhất thị trường.
  - **Interactive FX Converter 2.0**: Quy đổi 2 chiều USD $\leftrightarrow$ VND, tự động đề xuất ngân hàng có lợi nhất.
  - **Dual Chart Suite**: Biểu đồ cột phân cực và biểu đồ xu hướng đường toàn diện năm 2026 với bộ lọc nhanh (7N, 14N, 1T, 3T, 252N+).
  - **Keyboard Navigation**: Phím `[` / `]` đổi ngày, `T` đổi theme, `/` tìm kiếm nhanh.
- **Quy trình build & deploy**:
  - Cập nhật dữ liệu vào `docs/rates_history.json`.
  - Sinh mã HTML qua generator script và kiểm tra cú pháp AST bằng `node -c`.
  - Đẩy lên nhánh `main`, GitHub Actions tự động kích hoạt workflow `Deploy GitHub Pages`.

## 3. Tích Hợp Extension Phân Tích: `vscode-smart-data-viewer`
- Extension `vnstock.vscode-smart-data-viewer` (v0.6.0) được lưu trong `tools/` và cài đặt tại các thư mục extension của IDE.
- Cho phép mở trực tiếp `TyGia_Banking.xlsx`, `docs/rates_history.json`, hoặc các file CSV để chạy truy vấn DuckDB SQL và vẽ đồ thị EDA ngay trong trình soạn thảo.

## 4. Script Lịch Sử & Backfill
- `scripts/fetch_historical.py`: Backfill dữ liệu ngoại tệ VCB.
- `scripts/fetch_historical_all_banks.py`: Tải và chuẩn hóa lịch sử tỷ giá USD đa ngân hàng từ 01/01/2026 đến ngày hiện tại.

## 5. Lịch Trình Tự Động: `scripts/setup_schedule.ps1`
- Chạy trên PowerShell Administrator để đăng ký Windows Task Scheduler.
- Lịch chạy: **23h00 mỗi ngày**, tự động retry 3 lần, kích hoạt chế độ `StartWhenAvailable` để chạy bù nếu máy tính mở sau 23h.

---

# Quy Tắc Bất Biến (Constraints)
- Đường dẫn cố định dự án: `D:\Tygia-Tudong`
- Chuỗi thời gian luôn mở rộng động đến `datetime.now()` (Rolling Calendar), không hardcode ngày kết thúc trong quá khứ.
- Định dạng hiển thị số tiền/tỷ giá bắt buộc tuân theo hệ `,000,000 VND`.
- Tuyệt đối không sử dụng Playwright/trình duyệt nặng cho tác vụ cào hằng ngày; duy trì giải pháp thuần Python nhanh, nhẹ và ổn định.
