# SCADA Load Forecast Offline PWA LV5

Bản LV5 dùng để xây dựng quy trình dự báo phụ tải offline cho môi trường SCADA không có internet.

## Điểm chính

- Chạy offline, không CDN, không server ngoài, không `pwa.js`.
- Đọc dữ liệu: CSV/TXT/TSV/JSON và Excel `.xlsx/.xlsm` bằng thư viện nội bộ.
- Hiệu chỉnh dữ liệu trực tiếp trong bảng.
- Lọc theo một ngày, nhiều ngày, hoặc khoảng ngày để sửa nhanh.
- Điền nhanh nhiệt độ, mưa, ngày lễ, bất thường, cắt điện/sự cố, chuyển tải cho các dòng được chọn.
- Tự nhận dạng ngày nghỉ/lễ offline.
- Huấn luyện GBDT bằng JavaScript thuần.
- Xuất `model_gbdt.json` để đưa vào mạng SCADA.
- Nạp model và dự báo offline trong mạng SCADA.

## Nâng cấp LV5

LV5 bổ sung đầy đủ 8 nhóm chức năng:

1. **Kiểm tra chất lượng dữ liệu**
   - Thiếu/sai thời gian.
   - Thiếu/sai P.
   - P âm, P thấp bất thường.
   - Trùng mốc thời gian.
   - Mất mốc thời gian.
   - P tăng/giảm đột biến.
   - Xuất `quality_report.csv`.

2. **Nội suy / bổ sung dữ liệu mất mẫu**
   - Bổ sung mốc thời gian thiếu.
   - Nội suy P tuyến tính, giữ giá trị trước, cùng giờ hôm trước hoặc cùng giờ tuần trước.
   - Nội suy P trống/lỗi trong các dòng hiện có.
   - Tự thêm cột `du_lieu_noi_suy`, `ghi_chu_xu_ly`, `bo_khoi_huan_luyen`.

3. **Dự báo theo từng trạm/lộ riêng**
   - Chọn trạm/lộ để huấn luyện model riêng.
   - Có nút huấn luyện theo từng trạm/lộ để tạo bundle model.
   - Có nút dự báo tất cả trạm/lộ.

4. **Mô hình lai LV5**
   - GBDT.
   - Similar Day.
   - Cùng giờ tuần trước.
   - Xu hướng gần nhất.
   - Bù sai số gần nhất.
   - Cho phép chỉnh trọng số.

5. **Biểu đồ thực tế / dự báo / sai số**
   - Biểu đồ validation thực tế và dự báo.
   - Biểu đồ sai số.
   - Hiển thị Pmax dự báo, giờ Pmax, sai số lớn nhất và giờ sai số lớn nhất.

6. **Cảnh báo quá tải theo ngưỡng vận hành**
   - Nhập ngưỡng theo từng trạm/lộ.
   - Ví dụ: `E22.1,38,42`.
   - Tự đánh dấu `Bình thường`, `CẢNH BÁO`, `NGUY HIỂM` trong bảng dự báo.

7. **Lưu/nạp cấu hình offline**
   - Lưu cấu hình vào trình duyệt.
   - Xuất `config_lv5.json`.
   - Nạp `config_lv5.json`.
   - Lưu ánh xạ cột, ngày lễ, tham số huấn luyện, trọng số mô hình lai, ngưỡng cảnh báo.
   - Có nút **Ép cập nhật bản mới** để xóa cache PWA/Service Worker cũ khi copy bản mới vào cùng thư mục.

8. **Chế độ Mạng ngoài / SCADA**
   - Mạng ngoài: hiện đầy đủ hiệu chỉnh, kiểm tra, nội suy, huấn luyện.
   - SCADA: ẩn các khối hiệu chỉnh/huấn luyện để tập trung nạp model, dự báo và cảnh báo.

## Quy trình khuyến nghị

### Máy ngoài mạng SCADA

1. Mở app LV5.
2. Chọn chế độ **Mạng ngoài**.
3. Nạp dữ liệu lịch sử Excel/CSV.
4. Kiểm tra chất lượng dữ liệu.
5. Nội suy hoặc đánh dấu bất thường nếu cần.
6. Huấn luyện model GBDT hoặc huấn luyện theo từng trạm/lộ.
7. Xuất `model_gbdt.json`.
8. Xuất `config_lv5.json` nếu cần.

### Máy trong mạng SCADA

1. Mở app LV5.
2. Chọn chế độ **SCADA**.
3. Nạp dữ liệu mới nhất.
4. Nạp `model_gbdt.json`.
5. Nạp cấu hình/ngưỡng nếu có.
6. Bấm dự báo trạm/lộ đang chọn hoặc dự báo tất cả trạm/lộ.
7. Xuất `forecast.csv`.

## Chạy PWA đúng chuẩn

Nếu mở trực tiếp bằng `file://`, phần đọc file, sửa dữ liệu, huấn luyện và dự báo vẫn chạy. Để Service Worker và cache PWA hoạt động đúng chuẩn, nên chạy qua localhost hoặc web server nội bộ:

```bash
python -m http.server 8080
```

Sau đó mở:

```text
http://localhost:8080/index.html
```

## Lưu ý về Excel `.xls`

Bản này nhúng bộ đọc `.xlsx/.xlsm` nhẹ. File `.xls` nhị phân cũ nên lưu lại thành `.xlsx` hoặc CSV UTF-8 trước khi đưa vào app. Nếu cần đọc `.xls` trực tiếp, có thể thay bằng SheetJS chính thức `xlsx.full.min.js` nội bộ.

## LV5.2 - Tích hợp đường dẫn SheetJS full

Bản LV5.2 đã thêm sẵn đường dẫn nội bộ:

```text
libs/xlsx.full.min.js
```

`index.html` nạp thư viện theo thứ tự:

```html
<script src="libs/pako.min.js"></script>
<script src="libs/xlsx.full.min.js"></script>
<script src="libs/sheetjs-xlsx-lite.js"></script>
<script src="app.js"></script>
```

Cơ chế hoạt động:

- Nếu `libs/xlsx.full.min.js` là bản SheetJS full chính thức, ứng dụng sẽ dùng SheetJS full để đọc Excel.
- Nếu file này vẫn là placeholder, ứng dụng tự dùng `sheetjs-xlsx-lite.js` để đọc `.xlsx/.xlsm` cơ bản.
- File `.xls` nhị phân cũ cần thay placeholder bằng SheetJS full chính thức.

Nguồn khuyến nghị khi tải ở máy ngoài mạng SCADA:

```text
https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
```

Sau khi thay file, hãy bấm nút **Ép cập nhật bản mới** trong ứng dụng để xóa cache PWA cũ.


## LV5.2 - Sửa chức năng nội suy

Bản LV5.2 sửa phần nội suy theo thực tế vận hành:

- Nút **Nội suy P trống/lỗi trong dòng hiện có** xử lý cả P trống, P sai định dạng và P <= ngưỡng P thấp bất thường.
- Nếu đã chọn dòng rồi bấm **Bổ sung mốc thiếu**, app sẽ vừa bổ sung mốc thời gian thiếu, vừa tự nội suy P lỗi/trống/<= ngưỡng trong các dòng đã chọn.
- Phạm vi xử lý ưu tiên các dòng đã chọn; nếu chưa chọn dòng, dùng bộ lọc ngày/bảng hiện tại.
- Nội suy tôn trọng bộ lọc ngày đang dùng để tránh tạo mốc ngoài ngày đang hiệu chỉnh.

Ví dụ: nếu một mốc 21:00 có P = 0 và ngưỡng P thấp bất thường đặt là 0, hãy chọn dòng đó rồi bấm **Nội suy P trống/lỗi trong dòng hiện có** hoặc **Bổ sung mốc thiếu**. App sẽ lấy mốc hợp lệ trước/sau để nội suy và đánh dấu `du_lieu_noi_suy = 1`.


## LV5.3 - Xử lý P=0/P thấp theo cờ vận hành

Bản LV5.3 sửa logic theo thực tế vận hành:

- `P = 0` hoặc P thấp **không luôn luôn là lỗi đo**.
- Nếu dòng có cờ `cắt điện/sự cố`, `chuyển tải` hoặc `bất thường`, app mặc định **giữ số đo gốc** và có thể đánh dấu `bo_khoi_huan_luyen = 1` để model không học sai.
- Nếu P thấp không có cờ vận hành, app coi là dữ liệu cần xử lý: nội suy, đánh dấu bất thường hoặc kiểm tra lại.
- Thêm nút **Xử lý P=0/P thấp theo cờ vận hành**.
- Thêm lựa chọn **Cách xử lý P=0/P thấp**:
  - Tự động: có cờ vận hành thì giữ và bỏ huấn luyện.
  - Nội suy cả P thấp dù có cờ vận hành.
  - Không nội suy dòng có cờ vận hành, chỉ bỏ huấn luyện.
- Kiểm tra chất lượng có cảnh báo riêng cho trường hợp `chuyển tải` nhưng chưa thấy lộ/trạm khác tăng tải cùng thời điểm.

Khuyến nghị: với dữ liệu để huấn luyện dự báo, các dòng cắt điện/chuyển tải nên **bỏ khỏi huấn luyện** hoặc **nội suy thành phụ tải nền giả định**, không để model học trực tiếp giá trị 0 như một ngày bình thường.
