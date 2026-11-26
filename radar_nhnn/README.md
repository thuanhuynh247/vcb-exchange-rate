# Radar tiền tệ NHNN

Trang: `docs/radar/index.html` → https://thuanhuynh247.github.io/vcb-exchange-rate/docs/radar/

## Lấy số liệu từ đâu
Các trang biểu đồ của sbv.gov.vn gọi một API nội bộ của Liferay (không công bố nhưng mở cho khách vãng lai):

| Nguồn | Mã cấu trúc | Từ năm |
|---|---|---|
| Tỷ giá trung tâm VND/USD | 137473 | 2009 |
| Tỷ giá tham khảo tại Cục QLNH | 3450514 | ~2022 |
| Lãi suất & doanh số liên ngân hàng | 3450260 | 2014 |
| Lãi suất tái cấp vốn, tái chiết khấu | 3450482 | — |

`https://sbv.gov.vn/o/headless-delivery/v1.0/content-structures/<mã>/structured-contents?pageSize=200&sort=datePublished:desc`

OMO và CPI không có cấu trúc dữ liệu nên đọc từ HTML của trang `/vi/nghiệp-vụ-thị-trường-mở` và `/vi/cpi`.

## Giới hạn của máy chủ NHNN
Tường lửa trả **403** khi một địa chỉ IP gọi khoảng 5 lần liền nhau trong 1–2 phút. Vì vậy bộ cào:
- giãn 20 giây giữa hai lần gọi, tối đa 4 lần gọi mỗi lượt;
- chỉ gọi nguồn có khả năng đã có số mới (tỷ giá: khi chưa có số hôm nay; liên ngân hàng: 2 giờ/lần; CPI, lãi suất điều hành: 12 giờ/lần);
- gặp 403 thì dừng ngay lượt đó, lượt sau làm tiếp;
- tải lịch sử dần bằng các lần gọi còn dư (lùi theo ngày đăng, ghi vị trí vào `docs/radar/data/backfill.json`).

## Chạy
- Tự động: `.github/workflows/radar_nhnn.yml`, 30 phút/lần thứ 2–6 từ 7:00 đến 17:30 và 23:15 hằng ngày.
- Bằng tay: tab Actions → "Radar NHNN" → Run workflow (tích "Tải lại toàn bộ lịch sử" nếu cần làm lại từ đầu).
- Trên máy Windows (dự phòng nếu GitHub bị chặn lâu): `powershell -ExecutionPolicy Bypass -File radar_nhnn\cai_lich_windows.ps1`.
