# SCADA Load Forecast Offline PWA LV3

Công cụ HTML/PWA offline để tổng hợp dữ liệu phụ tải, hiệu chỉnh dữ liệu, huấn luyện mô hình GBDT và xuất `model_gbdt.json` để sử dụng trong mạng SCADA không có internet.

## Điểm mới LV3

- Đọc trực tiếp file Excel `.xlsx` và `.xlsm` trong trình duyệt.
- Giữ hỗ trợ CSV / TXT / TSV / JSON từ LV2.
- Có chọn sheet Excel sau khi nạp file.
- Nhúng nội bộ:
  - `libs/pako.min.js` để giải nén file `.xlsx`.
  - `libs/sheetjs-xlsx-lite.js` là lớp đọc Excel theo API kiểu SheetJS, chạy offline.
- Vẫn giữ chức năng sửa dữ liệu trực tiếp, lưu offline, xuất CSV/JSON đã hiệu chỉnh, huấn luyện GBDT, xuất `model_gbdt.json`.
- Có file mẫu `sample_load_data.xlsx` để kiểm thử đọc Excel.

## Lưu ý về SheetJS chính thức và file .xls

Bản LV3 hiện có bộ đọc `.xlsx/.xlsm` nhúng sẵn để chạy offline. File Excel `.xls` nhị phân cũ cần thư viện SheetJS chính thức `xlsx.full.min.js` hoặc cần lưu lại thành `.xlsx` / CSV UTF-8 trước khi nạp.

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

## Quy trình huấn luyện ngoài mạng SCADA

```text
1. Mở LV3 trên máy ngoài mạng SCADA.
2. Nạp dữ liệu lịch sử dạng .xlsx/.xlsm/.csv/.json.
3. Chọn sheet nếu là Excel.
4. Kiểm tra ánh xạ cột.
5. Hiệu chỉnh dữ liệu lỗi trong bảng nếu cần.
6. Bấm "Lưu thay đổi vào dữ liệu".
7. Huấn luyện GBDT.
8. Xuất model_gbdt.json.
```

## Quy trình dùng trong mạng SCADA

```text
1. Copy toàn bộ thư mục LV3 vào máy SCADA hoặc web server nội bộ.
2. Nạp dữ liệu phụ tải mới nhất.
3. Nạp model_gbdt.json đã huấn luyện ở bên ngoài.
4. Chọn trạm/lộ/khu vực.
5. Chạy dự báo offline.
6. Xuất forecast.csv nếu cần.
```

## Cột dữ liệu khuyến nghị

Tối thiểu:

```text
time, p_mw
```

Nên có thêm:

```text
station, temperature, rain, is_holiday, is_abnormal, is_outage, is_transfer
```

Ý nghĩa:

```text
time          : thời gian đo
station       : trạm / lộ / khu vực
p_mw          : công suất tác dụng MW
q_mvar        : công suất phản kháng nếu có
temperature   : nhiệt độ
rain          : mưa, có thể là 0/1 hoặc lượng mưa
is_holiday    : ngày nghỉ/lễ
is_abnormal   : dữ liệu bất thường
is_outage     : cắt điện/sự cố
is_transfer   : chuyển tải/kết lưới bất thường
```

Các dòng `is_abnormal`, `is_outage`, `is_transfer` sẽ được loại khỏi tập huấn luyện để tránh làm sai mô hình.

## Ghi chú vận hành

- Không dùng CDN.
- Không cần internet.
- Không có `pwa.js`.
- Service Worker cache các file nội bộ để dùng offline.
- File `.xlsx` lớn có thể mất thời gian đọc; nên lọc dữ liệu lịch sử theo trạm/lộ hoặc theo khoảng thời gian cần thiết.
