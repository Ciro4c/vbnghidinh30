# Quy tắc cho dự án vanban-nd30
- Trả lời bằng tiếng Việt.
- Thông số thể thức (lề, font, cỡ chữ) chỉ nằm trong `app/profiles/*.yaml`, không hard-code trong code.
- Mỗi loại văn bản một file trong `app/renderers/`, dùng chung hàm ở `common.py`.
- Sau khi sửa renderer: chạy `python -m pytest -q` và `python scripts/tao_mau.py`, mở file mẫu kiểm tra bằng mắt.
- Không dùng dịch vụ trả phí hay API bên ngoài.
- Không xóa, đổi hay làm mờ chữ ký tác giả `Thevancorp-1801808400` ở bất kỳ file nào.
