# Tool Test Tự Động — Dành cho người phát triển

- Chạy từ mã nguồn bằng `ChayTool.bat`, `start.bat` hoặc `python app.py`. Tool mở cửa sổ Edge dạng app tại `http://127.0.0.1:8620`.
- Thư viện cần cài: `pip install playwright openpyxl pillow`. **Không cần** chạy `playwright install` vì tool dùng Edge/Chrome có sẵn.
- Các file mã nguồn:

  | File | Nội dung |
  |---|---|
  | `scanner.py` | Chế độ quét từ URL: dò trang, sinh test case theo loại, chạy |
  | `scan_types.py` | Loại test, tùy chọn của từng loại, luật Business rule |
  | `engine.py` | Chạy theo Excel |
  | `excel_io.py` | Đọc Excel, tạo file mẫu |
  | `report.py` | Xuất Excel, file Test Scenario theo file mẫu (luồng, truy vết, tổng hợp), điền vào file Excel có sẵn, các bảng riêng, HTML |
  | `columns.py` | Danh mục cột, mẫu cột của công ty |
  | `app.py` | Server và API |
  | `web/` | Giao diện (`scan.js` là tab Test web: 6 bước) |
  | `demo/` | Web giả lập có 4 lỗi cố ý để tập |

- Lỗi cố ý trong `demo/`:
  - Số lượng nhận 101.
  - Điểm nhận -1 hoặc 11.
  - Viewer vẫn thấy nút Xóa.
  - Cookie thiếu HttpOnly.
- Dữ liệu lưu ở `%LOCALAPPDATA%\ToolTestTuDong`:
  - `settings.json`: cài đặt, **không lưu mật khẩu**.
  - `runs\`: kết quả các lần chạy.
  - `do-trang\`: kết quả dò trang theo URL.
  - `phien\`: phiên đăng nhập đã lưu.
  - `anh-chuan\`: ảnh chuẩn để so sánh.
- Test:
  - `python tests/test_excel.py`
  - `python tests/smoke_api.py <port>`: chế độ Excel.
  - `python tests/scan_demo.py <port>`: chế độ quét URL trên web demo.
  - `python tests/scenario_map.py <port>`: sinh / ghép Test Scenario trên web demo.
- Đóng gói: `build_exe.bat` → `dist\ToolTestTuDong\` (PyInstaller onedir, không kèm trình duyệt) và `dist\ToolTestTuDong_v<phiên bản>_Windows.zip` để đưa lên GitHub Releases. Đổi `VERSION` trong `app.py` trước khi build bản mới.
