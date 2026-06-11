# SCADA Load Forecast Offline PWA LV4

Công cụ HTML/PWA offline để tổng hợp dữ liệu phụ tải, hiệu chỉnh dữ liệu, huấn luyện mô hình GBDT và xuất `model_gbdt.json` để sử dụng trong mạng SCADA không có internet.

## Điểm mới LV4

- Thêm bộ lọc bảng hiệu chỉnh theo:
  - một ngày;
  - nhiều ngày nhập danh sách;
  - từ ngày đến ngày.
- Khi mỗi ngày có nhiều mốc dữ liệu, chọn một ngày sẽ hiển thị toàn bộ các mốc của ngày đó trên bảng hiệu chỉnh.
- Thêm chọn nhanh:
  - chọn tất cả dòng đang hiển thị;
  - chọn tất cả dòng sau bộ lọc ngày/bộ lọc lỗi;
  - bỏ chọn tất cả.
- Thêm điền nhanh cho các dòng được chọn:
  - nhiệt độ;
  - mưa;
  - ngày nghỉ/lễ;
  - bất thường;
  - cắt điện/sự cố;
  - chuyển tải;
  - trạm/lộ;
  - cột tùy chọn bất kỳ.
- Thêm tự nhận dạng ngày nghỉ/lễ offline theo quy tắc:
  - Thứ Bảy;
  - Chủ Nhật;
  - ngày lễ dương cố định: 01-01, 30-04, 01-05, 02-09;
  - danh sách ngày/khoảng nghỉ bổ sung do người dùng nhập.
- Việt hóa tiêu đề bảng hiển thị và có nút đổi tiêu đề cột sang tiếng Việt.

## Giữ chức năng từ LV3

- Đọc trực tiếp file Excel `.xlsx` và `.xlsm` trong trình duyệt.
- Giữ hỗ trợ CSV / TXT / TSV / JSON.
- Có chọn sheet Excel sau khi nạp file.
- Nhúng nội bộ:
  - `libs/pako.min.js` để giải nén file `.xlsx`.
  - `libs/sheetjs-xlsx-lite.js` là lớp đọc Excel theo API kiểu SheetJS, chạy offline.
- Sửa dữ liệu trực tiếp, lưu offline, xuất CSV/JSON đã hiệu chỉnh, huấn luyện GBDT, xuất `model_gbdt.json`.
- Có file mẫu `sample_load_data.xlsx` để kiểm thử đọc Excel.

## Lưu ý về SheetJS chính thức và file .xls

Bản LV4 hiện có bộ đọc `.xlsx/.xlsm` nhúng sẵn để chạy offline. File Excel `.xls` nhị phân cũ cần thư viện SheetJS chính thức `xlsx.full.min.js` hoặc cần lưu lại thành `.xlsx` / CSV UTF-8 trước khi nạp.

Trong vận hành thực tế, nên ưu tiên xuất từ Excel/SCADA thành `.xlsx` hoặc CSV UTF-8.

## Cấu trúc file

```text
index.html
app.js
sw.js
manifest.webmanifest
README.md
sample_load_data.csv
sample_load_data.xlsx
libs/pako.min.js
libs/sheetjs-xlsx-lite.js
```

## Cách chạy

Mở trực tiếp `index.html` vẫn dùng được phần đọc file, sửa dữ liệu, huấn luyện và dự báo.

Để PWA cache offline hoạt động đúng chuẩn, chạy qua localhost hoặc web server nội bộ:

```bash
python -m http.server 8080
```

Sau đó mở:

```text
http://localhost:8080/index.html
```

## Quy trình hiệu chỉnh dữ liệu theo ngày

```text
1. Nạp file dữ liệu.
2. Kiểm tra ánh xạ cột.
3. Ở phần Hiệu chỉnh dữ liệu, chọn kiểu lọc ngày:
   - Một ngày;
   - Nhiều ngày;
   - Từ ngày đến ngày.
4. Bấm chọn tất cả dòng đang hiển thị hoặc chọn tất cả dòng sau lọc.
5. Nhập giá trị điền nhanh, ví dụ nhiệt độ 36.5 hoặc cờ ngày nghỉ/lễ = 1.
6. Bấm “Điền nhanh vào dòng đã chọn”.
7. Bấm “Lưu thay đổi vào dữ liệu”.
```

## Tự nhận dạng ngày nghỉ/lễ

Công cụ tự nhận dạng theo quy tắc offline, không cần mạng.

Mặc định gồm:

```text
Thứ Bảy
Chủ Nhật
01-01
30-04
01-05
02-09
```

Với Tết âm lịch, nghỉ bù hoặc lịch riêng của đơn vị, nhập thêm vào ô “Ngày/khoảng nghỉ bổ sung”, ví dụ:

```text
2026-02-16..2026-02-20, 2026-04-18
```

Sau đó bấm:

```text
Tự nhận dạng cho dòng đã chọn
hoặc
Tự nhận dạng cho toàn bộ dòng sau lọc
```

## Quy trình huấn luyện ngoài mạng SCADA

```text
1. Mở LV4 trên máy ngoài mạng SCADA.
2. Nạp dữ liệu lịch sử dạng .xlsx/.xlsm/.csv/.json.
3. Chọn sheet nếu là Excel.
4. Kiểm tra ánh xạ cột.
5. Hiệu chỉnh dữ liệu lỗi, nhiệt độ, ngày nghỉ/lễ, bất thường nếu cần.
6. Bấm “Lưu thay đổi vào dữ liệu”.
7. Huấn luyện GBDT.
8. Xuất model_gbdt.json.
```

## Quy trình dùng trong mạng SCADA

```text
1. Copy toàn bộ thư mục LV4 vào máy SCADA hoặc web server nội bộ.
2. Nạp dữ liệu phụ tải mới nhất.
3. Nạp model_gbdt.json đã huấn luyện ở bên ngoài.
4. Chọn trạm/lộ/khu vực.
5. Chạy dự báo offline.
6. Xuất forecast.csv nếu cần.
```

## Cột dữ liệu khuyến nghị

Tối thiểu:

```text
Thời gian, P_MW
```

Nên có thêm:

```text
Trạm/Lộ, Nhiệt độ, Mưa, Ngày nghỉ/lễ, Bất thường, Cắt điện/Sự cố, Chuyển tải
```

Các dòng `Bất thường`, `Cắt điện/Sự cố`, `Chuyển tải` = 1 sẽ được loại khỏi tập huấn luyện để tránh làm sai mô hình.

## Ghi chú vận hành

- Không dùng CDN.
- Không cần internet.
- Không có `pwa.js`.
- Service Worker cache các file nội bộ để dùng offline.
- Bộ lọc ngày chỉ giới hạn bảng hiệu chỉnh, không xóa dữ liệu gốc.
