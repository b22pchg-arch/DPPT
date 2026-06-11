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
