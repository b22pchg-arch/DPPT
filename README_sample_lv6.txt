Dữ liệu mẫu LV6 theo quy tắc chỉ danh Đơn vị/Trạm/Lộ_Đơn vị nối vòng/Trạm nối vòng/Lộ nối vòng

File chính:
- sample_load_data_lv6_chidanh.csv
- sample_load_data_lv6_chidanh.xlsx

Các chỉ danh mẫu:
1) Tuyên Quang/E22.1/473_Hà Giang/E14.1/471
2) Hà Giang/E14.1/471_Tuyên Quang/E22.1/473
3) Hà Giang/E14.2/473
4) Tuyên Quang/E22.2/475
5) Tuyên Quang/E22.3/477_Hà Giang/E14.3/475
6) Hà Giang/E14.3/475_Tuyên Quang/E22.3/477

Các tình huống test:
- 2026-04-28 14:00..15:30: chuyển tải 473 -> 471, lộ chính P=0 và lộ nối vòng tăng tương ứng.
- 2026-04-29 10:00..10:45: cắt điện/sự cố thật tại Hà Giang/E14.2/473, P=0 có cờ cắt điện.
- 2026-04-30 02:15: P=0 không có cờ vận hành tại Tuyên Quang/E22.2/475, cần phát hiện là lỗi dữ liệu.
- 2026-05-02 19:00: phụ tải tăng đột biến chưa đánh dấu tại Tuyên Quang/E22.3/477_Hà Giang/E14.3/475.
- 2026-04-27 03:00 và 03:15: thiếu mốc thời gian của Hà Giang/E14.3/475_Tuyên Quang/E22.3/477.

Gợi ý test trong LV6:
1. Nạp file sample_load_data_lv6_chidanh.xlsx hoặc .csv.
2. Bấm "Tách chỉ danh LV6 vào cột riêng".
3. Xuất designation_map.csv để kiểm tra đơn vị/trạm/lộ/nối vòng.
4. Bấm "Kiểm tra dữ liệu".
5. Bấm "Phân tích chuyển tải / P=0".
6. Lọc ngày 2026-04-28 để kiểm tra chuyển tải chéo.
7. Lọc ngày 2026-04-30 để kiểm tra P=0 không có cờ vận hành.
8. Lọc ngày 2026-04-27 và chọn Hà Giang/E14.3/475_Tuyên Quang/E22.3/477 để test bổ sung mốc thiếu.
