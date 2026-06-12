# SCADA Load Forecast Offline PWA LV9.8.1

Bản LV9.8.1 phát triển từ lõi ổn định LV8.9/LV8.6, bổ sung giao diện **Workflow / Workbench** để người vận hành thao tác theo từng luồng, chỉ hiển thị bước đang làm thay vì toàn bộ mục dọc.

## Điểm mới LV9.8.1

- Chọn nhanh 3 luồng công việc:
  - **Luồng A - Tạo model ngoài SCADA**: nạp dữ liệu lịch sử, ánh xạ, tách chỉ danh, kiểm tra chất lượng, nội suy, phân tích vận hành, huấn luyện, so sánh chiến lược và xuất model.
  - **Luồng B - Dự báo trong mạng SCADA**: nạp dữ liệu mới, nạp model vận hành, dự báo nhanh/đa cấp, báo cáo và cập nhật forecast vào RAM.
  - **Luồng C - Đánh giá sai số và hiệu chỉnh model**: nạp dữ liệu thực tế + dữ liệu dự báo, đánh giá sai số, tạo hiệu chỉnh, áp dụng vào model và xuất model hiệu chỉnh.
- Mỗi luồng có thanh bước riêng, nút **Quay lại / Tiếp tục / Thực hiện chính**. Từ LV9.8.1, khung chức năng của bước hiện tại được đưa trực tiếp vào vùng làm việc bên phải, không còn nằm dưới shell workflow.
- Mặc định chỉ hiển thị khung chức năng của bước đang thao tác để tăng không gian cho bảng và giảm khựng.
- Có **Chế độ chuyên gia / Hiển thị tất cả mục** để quay về giao diện dọc đầy đủ như LV8.9.
- Dashboard và nhật ký thao tác vẫn là quan sát/ghi nhận, **không khóa nút**, không can thiệp lõi tính toán.

## Nguyên tắc vận hành

LV9.8.1 không thay đổi thuật toán lõi của LV8.9. Giao diện Workflow chỉ ẩn/hiện các khung chức năng phù hợp với bước hiện tại. Nếu cần truy cập mọi nút cùng lúc, bật **Hiển thị tất cả mục**.

## Cách dùng nhanh

1. Mở ứng dụng và chọn Luồng A/S/B/C.
2. Bấm từng bước ở thanh bên trái hoặc dùng **Quay lại / Tiếp tục**.
3. Ở mỗi bước, đọc hướng dẫn trong khung LV9.8.1 rồi thao tác trên khung chức năng đang hiển thị.
4. Có thể bấm **Thực hiện chính** để kích hoạt nút chính của bước hiện tại.
5. Khi cần kiểm tra toàn bộ, bật **Chế độ chuyên gia**.

## Cập nhật PWA

Sau khi chép bản mới vào máy, bấm **Ép cập nhật bản mới**, sau đó đóng/mở lại ứng dụng để chắc chắn cache cũ đã được thay bằng LV9.8.1.

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


## LV9.8.1.1 - Phân luồng thật

Bản LV9.8.1 ban đầu chỉ thêm thanh workflow nhưng thiếu CSS ẩn các khu vực cũ nên giao diện vẫn giống dạng nhiều mục dọc. LV9.8.1.1 sửa lại theo đúng thiết kế Workbench:

- Mặc định mở màn hình chọn Luồng A/S/B/C.
- Sau khi chọn luồng, chỉ các khung của bước hiện tại được hiển thị.
- Các mục không liên quan được ẩn hoàn toàn để tăng không gian thao tác.
- Chế độ chuyên gia vẫn cho phép xem toàn bộ giao diện cũ.
- Dashboard/Nhật ký chỉ là ngăn phụ theo dõi, không khóa nút và không can thiệp chức năng.
- Lõi tính toán giữ ổn định từ LV8.9/LV8.6.


## LV9.8.1 - Workbench đúng nghĩa

LV9.1 đã phân luồng thật nhưng khung thao tác vẫn nằm bên dưới shell workflow, làm người vận hành phải kéo xuống và mất không gian. LV9.8.1 sửa bố cục:

- Khung chức năng của bước hiện tại được **đưa trực tiếp vào vùng làm việc bên phải** của Workbench.
- Thanh bước bên trái giữ cố định, vùng thao tác bên phải chiếm toàn bộ phần còn lại.
- Các card khác vẫn bị ẩn khi ở Workflow; khi bật Chế độ chuyên gia, mọi card được trả về đúng vị trí cũ.
- Không khóa nút, không thay đổi thuật toán lõi LV8.9/LV8.6.
- Service Worker cache đổi sang LV9.8.1 và cache đúng file `workflow_lv9.js`.


## LV9.8.1 - Tối ưu không gian Workbench

- Mặc định dùng thanh điều khiển tối giản ở phía trên: chọn luồng, chọn bước, Quay lại, Thực hiện, Tiếp tục.
- Ẩn phần mô tả dài, thanh luồng lớn và bảng bước bên trái để khu vực chức năng chiếm gần toàn bộ màn hình.
- Nút **Bước** mở/ẩn danh sách bước dạng hẹp khi cần nhảy nhanh.
- Nút **Đầy đủ / Tối giản** chuyển giữa giao diện hướng dẫn đầy đủ và giao diện thao tác gọn.
- Không khóa nút, không thay đổi lõi tính toán ổn định của LV8.6/LV8.9.

## LV9.8.1 - Tách phần dùng chung của Luồng B và Luồng C

Bản LV9.8.1 xử lý nhược điểm Luồng B và Luồng C có nhiều bước giống nhau bằng cách thêm **Luồng S - Chuẩn bị chung cho B/C**.

Luồng mới:

- **Luồng A**: tạo model ngoài mạng SCADA.
- **Luồng S**: nạp dữ liệu, ánh xạ cột, tách chỉ danh và nạp model vận hành. Đây là phần dùng chung cho dự báo và đánh giá.
- **Luồng B**: chỉ tập trung dự báo vận hành, báo cáo, cảnh báo và cập nhật forecast vào RAM.
- **Luồng C**: chỉ tập trung đánh giá sai số, tạo hiệu chỉnh, áp dụng hiệu chỉnh và xuất model/hồ sơ hiệu chỉnh.

Luồng thao tác khuyến nghị:

1. Tạo model: dùng Luồng A.
2. Dự báo: dùng Luồng S để chuẩn bị dữ liệu/model, sau đó chuyển sang Luồng B.
3. Đánh giá/hiệu chỉnh: dùng Luồng S nếu cần nạp dữ liệu/model, sau đó chuyển sang Luồng C.

Việc này giúp giảm lặp bước giữa B và C, đồng thời người vận hành hiểu rõ: **S là chuẩn bị chung, B là dự báo, C là đánh giá/hiệu chỉnh**.


## LV9.8.1 - Tách Mục 10 theo bước Workflow

LV9.8.1 tách khối Mục 10 cũ thành nhiều khung riêng để phù hợp với từng bước của luồng công việc:

- 10A: Nạp model vận hành
- 10B: Dự báo nhanh bằng model vận hành
- 10C: Dự báo đa cấp LV8.5
- 10D: Đánh giá sai số từ dữ liệu thực tế
- 10E: Hiệu chỉnh mô hình từ sai số
- 10F: So sánh trước/sau hiệu chỉnh
- 10G: Ngưỡng cảnh báo vận hành

Trong Workflow, Luồng S chỉ hiện 10A khi nạp model; Luồng B chỉ hiện 10B/10C/10G khi dự báo; Luồng C chỉ hiện 10D/10E/10F khi đánh giá và hiệu chỉnh. Lõi tính toán không thay đổi so với LV8.9 ổn định.


## LV9.8.1 - Làm rõ Luồng B2/B3/B4/B5

LV9.8.1 tách khối dự báo đa cấp 10C thành bốn khung nhỏ để người vận hành không bị lẫn thao tác:

```text
10C1 / B2: Chạy dự báo hoặc tạo báo cáo đa cấp LV8.5
10C2 / B3: Xem kết quả Pmax, giờ Pmax, MWh, phụ tải theo ca và cảnh báo
10C3 / B4: Cập nhật forecast vào RAM để dự báo nối tiếp
10C4 / B5: Xuất forecast_summary_lv8_5.csv
```

Các nút trong 10C đã được đổi nhãn theo bước B2/B4/B5 để người dùng biết rõ đang thao tác ở bước nào. Luồng B khuyến nghị:

```text
B2 → bấm Dự báo LV8.5 hoặc Tạo báo cáo đa cấp
B3 → xem kết quả/cảnh báo, chỉnh ngưỡng nếu cần
B4 → cập nhật forecast vào RAM nếu muốn dự báo nối tiếp
B2 → chạy dự báo tiếp sau khi đã cập nhật RAM
B5 → xuất forecast_summary_lv8_5.csv
```

Lõi tính toán không đổi so với LV9.5/LV8.9; thay đổi chỉ là bố trí lại giao diện và nhãn nút để phù hợp quy trình vận hành.


## Bổ sung LV9.8.1

- Sửa B1: khi bấm **Dự báo trạm/lộ đang chọn**, ứng dụng hiển thị đồng thời biểu đồ và bảng kết quả forecast.
- Sửa B1: khi bấm **Dự báo tất cả trạm/lộ**, ứng dụng hiển thị bảng forecast của tất cả trạm/lộ và biểu đồ tổng P dự báo theo thời gian.
- Luồng B1 trong Workflow tự đưa thêm khung **Bảng dữ liệu chuẩn hóa / Dự báo** vào vùng làm việc, không phải chuyển sang chế độ chuyên gia để xem bảng.


## LV9.8.1 - Sửa B2 bị ẩn Mục 10B

Bản LV9.8.1 sửa lỗi trong Workbench Luồng B: khi vào B2, khung 10B bị ẩn nên người vận hành không thấy các nút **Dự báo trạm/lộ đang chọn** và **Dự báo tất cả trạm/lộ**.

Thay đổi:

- B2 hiển thị đồng thời **10B - Dự báo nhanh** và **10C1 - Dự báo đa cấp**.
- Có thể bấm dự báo nhanh theo trạm/lộ, dự báo tất cả trạm/lộ, hoặc chạy dự báo/tạo báo cáo đa cấp ngay trong cùng bước B2.
- Giữ bảng kết quả và biểu đồ trong vùng làm việc để kiểm tra ngay sau khi bấm.
- Không thay đổi lõi tính toán, chỉ sửa hiển thị Workflow.


## Ghi chú LV9.8.1

Sửa lỗi nhận diện khung `10B) B1 - Dự báo nhanh bằng model vận hành` trong Workflow: trước đó biểu thức nhận diện chỉ khớp `10B) Dự báo nhanh`, nên B1 có thể không đưa khung 10B vào Workbench. LV9.8.1 nhận diện đúng 10B ở B1 và vẫn giữ 10B trong B2.
