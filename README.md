# SCADA Load Forecast Offline PWA LV9

Bản LV9 phát triển từ lõi ổn định LV8.9/LV8.6, bổ sung giao diện **Workflow / Workbench** để người vận hành thao tác theo từng luồng, chỉ hiển thị bước đang làm thay vì toàn bộ mục dọc.

## Điểm mới LV9

- Chọn nhanh 3 luồng công việc:
  - **Luồng A - Tạo model ngoài SCADA**: nạp dữ liệu lịch sử, ánh xạ, tách chỉ danh, kiểm tra chất lượng, nội suy, phân tích vận hành, huấn luyện, so sánh chiến lược và xuất model.
  - **Luồng B - Dự báo trong mạng SCADA**: nạp dữ liệu mới, nạp model vận hành, dự báo nhanh/đa cấp, báo cáo và cập nhật forecast vào RAM.
  - **Luồng C - Đánh giá sai số và hiệu chỉnh model**: nạp dữ liệu thực tế + dữ liệu dự báo, đánh giá sai số, tạo hiệu chỉnh, áp dụng vào model và xuất model hiệu chỉnh.
- Mỗi luồng có thanh bước riêng, nút **Quay lại / Tiếp tục / Thực hiện chính**.
- Mặc định chỉ hiển thị khung chức năng của bước đang thao tác để tăng không gian cho bảng và giảm khựng.
- Có **Chế độ chuyên gia / Hiển thị tất cả mục** để quay về giao diện dọc đầy đủ như LV8.9.
- Dashboard và nhật ký thao tác vẫn là quan sát/ghi nhận, **không khóa nút**, không can thiệp lõi tính toán.

## Nguyên tắc vận hành

LV9 không thay đổi thuật toán lõi của LV8.9. Giao diện Workflow chỉ ẩn/hiện các khung chức năng phù hợp với bước hiện tại. Nếu cần truy cập mọi nút cùng lúc, bật **Hiển thị tất cả mục**.

## Cách dùng nhanh

1. Mở ứng dụng và chọn Luồng A/B/C.
2. Bấm từng bước ở thanh bên trái hoặc dùng **Quay lại / Tiếp tục**.
3. Ở mỗi bước, đọc hướng dẫn trong khung LV9 rồi thao tác trên khung chức năng đang hiển thị.
4. Có thể bấm **Thực hiện chính** để kích hoạt nút chính của bước hiện tại.
5. Khi cần kiểm tra toàn bộ, bật **Chế độ chuyên gia**.

## Cập nhật PWA

Sau khi chép bản mới vào máy, bấm **Ép cập nhật bản mới**, sau đó đóng/mở lại ứng dụng để chắc chắn cache cũ đã được thay bằng LV9.

## Thành phần chính

```text
index.html
app.js
workflow_lv9.js
sw.js
manifest.webmanifest
libs/
icons/
sample_load_data_lv6_chidanh.csv
sample_load_data_lv6_chidanh.xlsx
```

Không dùng CDN, không cần internet, không dùng `pwa.js`.
