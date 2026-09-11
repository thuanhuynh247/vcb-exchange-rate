# Financial Dashboard & FX Crawler Workspace Invariants

Tài liệu này định nghĩa các quy chuẩn bất biến bắt buộc phải tuân thủ khi phát triển, mở rộng và bảo trì hệ thống tại `D:\Tygia-Tudong`:

## 1. Định Dạng Số Tiền & Tỷ Giá (Financial Number Formatting)
- Bắt buộc sử dụng dấu phẩy `,` làm ký tự phân tách hàng nghìn theo chuẩn tài chính quốc tế (Hệ `,000,000 VND`).
- Không sử dụng dấu chấm `.` làm phân tách hàng nghìn cho tiền tệ/tỷ giá trên giao diện web hoặc xuất dữ liệu để tránh xung đột với các phần mềm bảng tính.
- Hàm format chuẩn: `Number(val).toLocaleString('en-US', { maximumFractionDigits: 2 })`. Áp dụng đồng bộ cho các thẻ KPI, bộ quy đổi ngoại tệ, bảng ma trận 15 ngân hàng, tooltip và trục đồ thị.

## 2. Trải Nghiệm Sao Chép Dữ Liệu (Copy-to-Clipboard UX)
- Mọi ô dữ liệu tỷ giá (Mua TM, Mua CK, Bán, Spread, Dòng bình quân), thẻ KPI và kết quả tính đổi ngoại tệ đều phải hỗ trợ sao chép 1-click.
- Luôn hiển thị Toast Notification phản hồi ngay khi sao chép thành công.
- Cung cấp sẵn các định dạng xuất tiện ích: Tab-separated (TSV) để dán trực tiếp vào Excel / Google Sheets và văn bản tóm tắt để gửi Zalo / Telegram.
- Cơ chế copy phải hỗ trợ cả `navigator.clipboard` và fallback `document.execCommand('copy')` để hoạt động an toàn trên mọi trình duyệt.

## 3. Chuỗi Thời Gian Động (Rolling Calendar Integrity)
- Không bao giờ hardcode mốc ngày kết thúc cố định trong quá khứ khi backfill hay tạo dữ liệu.
- Mọi tác vụ tổng hợp lịch sử hoặc cào dữ liệu phải tự động mở rộng đến `datetime.now()` để đảm bảo ngày gần nhất (hôm qua, hôm nay) luôn có mặt trong hệ thống và hiển thị ở đầu danh sách.

## 4. Kiểm Thử Cú Pháp Bắt Buộc Trước Khi Deploy (AST Validation)
- Khi sinh mã HTML/JS từ Python template string: không đặt ký tự xuống dòng trực tiếp trong string literal của JavaScript; sử dụng `String.fromCharCode(10)` hoặc `encodeURIComponent`.
- Luôn chạy lệnh kiểm tra cú pháp AST `node -c <file>` và chạy thử nghiệm runtime giả lập DOM trước khi commit mã nguồn lên git.
