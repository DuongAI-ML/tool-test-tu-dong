# Tool Test Tự Động

Bạn chỉ cần **nhập URL web**. Tool tự dò trang, bạn chọn làm Test Scenario hay Test Case, chọn cột và nhóm test, tool tự sinh, tự chạy trên Edge/Chrome thật của máy rồi xuất Excel theo mẫu cột của công ty. Ngoài ra vẫn chạy được theo file test case Excel viết tay.

## Tải về và chạy (Windows 10/11)

Máy cần có **Microsoft Edge** (Windows nào cũng có sẵn) hoặc Google Chrome. Chọn 1 trong 2 cách:

### Cách 1 — Bản EXE, không cần cài gì (khuyên dùng)

1. Vào **[Releases](https://github.com/DuongAI-ML/tool-test-tu-dong/releases/latest)**, tải file `ToolTestTuDong_v…_Windows.zip` (khoảng 57 MB).
2. Chuột phải file zip → **Extract All** (giải nén). Đừng chạy thẳng trong file zip.
3. Mở thư mục `ToolTestTuDong` → bấm đúp **ToolTestTuDong.exe**.
   Nếu Windows hiện "Windows protected your PC": bấm **More info** → **Run anyway** (tool chưa có chữ ký số, không phải virus).

### Cách 2 — Chạy từ mã nguồn (nhẹ để tải, dễ cập nhật)

1. Bấm nút xanh **Code → Download ZIP** ở trang này rồi giải nén. Hoặc nếu máy có git:
   ```bat
   git clone https://github.com/DuongAI-ML/tool-test-tu-dong.git
   ```
2. Bấm đúp **ChayTool.bat**.
   - Lần đầu: tool tự cài Python (nếu máy chưa có, qua winget) và thư viện (~45 MB) vào `%LOCALAPPDATA%\ToolTestTuDong\venv`. Mất khoảng 1–3 phút, cần mạng.
   - Các lần sau: mở tool ngay.
3. Cập nhật bản mới: `git pull` (hoặc tải ZIP mới) rồi bấm lại ChayTool.bat.

Không cần lệnh `playwright install`: tool điều khiển Edge / Chrome có sẵn trên máy.

## Dành cho người dùng (tester)

### Tab 🌐 Test web (cách dùng chính) — 1 trang, làm từ trên xuống

1. **Bạn muốn làm gì**: bấm thẻ **Test Scenario** hoặc **Test Case**.
2. **Cột**: hiện các cột của loại vừa chọn. Sửa tên cột, kéo ⋮⋮ hoặc bấm ‹ › để đổi thứ tự, ✕ để bỏ, "＋" để thêm. Bấm **Lưu mẫu cột** / **Lưu thành mẫu mới…**.
   - Test Scenario mặc định: ID | Nhóm | Test Scenario | AC liên quan | Priority.
   - Test Case mặc định (mẫu công ty): Mã TC, Chức năng, Requirement / Scenario, Tên TC, Preconditions, Environment, Steps, Kết quả mong đợi, Kết quả thực tế, Data ID, Test data, Priority, Kết quả test (Lần 1/2/3), Trạng thái Test tổng, Bug #.
3. **Nguồn**: nhập URL, chọn cách đăng nhập, bấm **🔍 Dò trang**. Với Test Case có thêm lựa chọn **Lấy Test Scenario từ file Excel có sẵn** để sinh test case từ scenario của bạn.
   - Dò ra quá ít nút / form, gặp captcha, trang chặn tool hoặc cần đăng nhập → tool hiện khung ⚠ nói rõ **lý do** và **cách xử lý** (Dò lại – chờ lâu hơn; Tôi tự đăng nhập / vượt captcha rồi dò lại; tắt Chạy ẩn).
4. **Nhóm test**: tick validation, functional, integration, business rule, permission, state, usability, UI, responsive, performance, security… Tùy chọn của từng nhóm nằm trong "Tùy chọn ▸". Ô **"Thêm scenario gợi ý cần test tay"**: tool thêm các scenario theo tính năng dò thấy (bộ lọc, Làm mới, Lưu nháp, Gửi phê duyệt, Xóa, Nhập file…) mà tool chưa tự kiểm tra được.
5. **Xem trước**: bảng đúng các cột đã chọn.
   - Test Scenario: chia theo **LUỒNG** như file mẫu (truy cập & giao diện, hiển thị danh sách, tìm kiếm & bộ lọc, phối hợp chức năng danh sách, thêm mới, lưu & phê duyệt, xem chi tiết, hành động trên bản ghi, nhập dữ liệu từ file). ID đánh liên tục TS_DHTD_001… theo các dòng đã tick. Bấm vào ô Test Scenario / AC liên quan để sửa, chọn Priority.
   - Test Case: nhóm theo scenario, có Steps / Kết quả mong đợi dự kiến. Chạy xong thì hiện kết quả thật.
6. **Xuất file**: đặt tên file (biến `{man_hinh}` `{ngay}` `{gio}` `{du_an}` `{nguoi_test}`), thư mục, rồi bấm:
   - **📄 Sinh file Excel mới**. File Test Scenario theo đúng file mẫu `4_TestScenario_DonHangTuyenDung.xlsx`: sheet "Test Scenario" (dòng 1 tiêu đề, dòng 2 ghi chú, dòng 4 header xanh 1F4E78, dòng LUỒNG gộp ô nền DDEBF7, độ rộng 14/13/95/20/10, freeze A5, Priority tô màu) + sheet "Ma tran truy vet" + "Tong hop".
   - **📥 Điền vào file Excel sẵn có**: chọn file (nút 📂, dán đường dẫn hoặc kéo-thả) → chọn sheet (tool nhận cột theo **tên tiêu đề**) → **Ghi tiếp** (ID nối tiếp file, VD từ TS_DHTD_044) hoặc **Thay toàn bộ** → lưu **bản sao** (mặc định) hoặc ghi đè file gốc (tool hỏi lại). Định dạng chép từ dòng có sẵn.
   - **▶ Chạy test** (Test Case): chạy các test case đã tick. Xong thì "Sinh file mới" kèm kết quả, hoặc "Điền vào file sẵn có" để ghi **Kết quả test (Lần N)**, Kết quả thực tế, Trạng thái tổng theo Mã TC.
   - Mã TC ở bảng xem trước, file sinh ra và lần chạy luôn trùng nhau.

Nút **Cấu hình thêm** ở bước 4: phạm vi dò, trình duyệt, chạy ẩn, thời gian chờ trang SPA; Business rule; ma trận phân quyền; luồng trạng thái; nút cho bấm thử.

File ví dụ: `vi-du/TestScenario_DHTD.xlsx`, gồm 21 scenario từ ảnh tester gửi.

### Chi tiết từng nhóm (trước đây là tab Quét từ URL)

Tab này làm theo đúng cách tester làm việc thật:

1. **Trang web cần test**
   - Nhập URL.
   - Chọn cách đăng nhập: không cần, tự đăng nhập bằng tay trong trình duyệt (tool chờ, dùng được cho OTP/captcha), hoặc để tool tự điền tài khoản.
   - Chọn phạm vi dò.
   - Bấm **🔍 Dò trang**.
2. **Những gì tool dò được**
   - Tool liệt kê đúng các trang, nút, popup, form, trường (có dấu * bắt buộc, độ dài, min/max) và các giá trị trạng thái đang có trên trang.
   - Tick nút nào thì tool được phép bấm thử nút đó. Nên bỏ tick các nút làm thay đổi dữ liệu thật.
3. **Chọn loại test**
   - Chọn từng loại, hoặc dùng gói "7 loại chính". Mỗi loại có dropdown cấu hình riêng:
     - **Validation**: bỏ trống ô (*), sai định dạng email/SĐT, vượt độ dài, số ngoài khoảng, tải tệp sai định dạng.
     - **Functional**: trang mở được, link, nút, tìm kiếm, phân trang, sắp xếp cột, các nút của form (Lưu, Lưu nháp, Gửi phê duyệt, Hủy).
     - **Integration**: dropdown bị trống, API lỗi, dữ liệu khớp API, biểu tượng "đang tải" bị treo.
     - **Business rule**: chọn **trường (lấy từ dữ liệu dò được)** và **luật** (bắt buộc, độ dài, khoảng số, email, SĐT, ngày, trường tự tính…). Tool tự sinh ca sai, ca biên và ca đúng.
     - **Permission**: dùng **ma trận phân quyền**, với chức năng và nút lấy từ dữ liệu dò được, nhân với vai trò của tài khoản 2. Mỗi ô chọn Được thấy hoặc Không được thấy.
     - **State**: mở/đóng bộ lọc, menu, toggle, popup. Bảng **chuyển trạng thái**: chọn form và nút, ví dụ bấm "Lưu nháp" thì phải ra trạng thái "Nháp".
     - **Usability**: nút có bị che khuất không, nút quá nhỏ, nút gạt có bấm được không, chữ quá nhỏ.
     - Các loại bổ sung: UI, Responsive, Performance, Security, Accessibility, Compatibility.
4. Bấm **▶ Chạy test**.
   - Case tool không chắc chắn sẽ **dừng lại hỏi bạn** Pass/Fail.
   - Tool không bao giờ tự ghi Pass cho case chưa được kiểm tra.

### Tab 📊 Kết quả & Xuất file

- Đặt tên file bằng các biến `{du_an}` `{ngay}` `{gio}` `{moi_truong}` `{lan_chay}` `{nguoi_test}`.
- Chọn ghi kết quả vào cột Lần 1, Lần 2 hoặc Lần 3.
- Bấm **Xuất**. File xuất ra gồm:
  - Sheet **TestCase** theo **"Mẫu công ty (theo file tester)"**:
    - Các cột: Mã TC, Chức năng, Requirement / Scenario, Tên TC, Preconditions, Environment, Steps, Kết quả mong đợi, Kết quả thực tế, Data ID, Test data, Priority, Kết quả test (Lần 1/2/3), Trạng thái Test tổng, Bug # (ảnh lỗi).
    - Các bước gộp trong 1 ô, có dòng nhóm theo loại test.
  - Sheet **Scenario**: ID, Nhóm, Test Scenario, AC liên quan, Priority.
  - **Mỗi loại test có 1 bảng riêng** như tester thường dùng: Ma trận Validation, Bảng quyết định (Business rule), Ma trận phân quyền, Bảng chuyển trạng thái, Kiểm tra tích hợp, Kiểm tra Usability, và các Checklist khác. Các bảng nối với sheet TestCase qua Mã TC.

### Các tab khác

- **Tab ☑ Cấu hình cột**: tick, đổi tên, đổi thứ tự cột; chọn kiểu ghi bước (gộp 1 ô hoặc mỗi bước 1 dòng); bật/tắt dòng nhóm và bảng riêng theo loại test.
- **Tab 📄 Chạy theo file Excel**: chạy test case viết sẵn trong Excel. Đọc được file thật của tester:
  - các bước "1. … 2. …" trong 1 ô,
  - dòng nhóm kiểu "LUỒNG NGHIỆP VỤ",
  - cột Lần 1/2/3.

  Có ô **URL web cần test** để đổi môi trường nhanh.

## Dành cho người phát triển

- Chạy từ mã nguồn bằng `start.bat` hoặc `python app.py`. Tool mở cửa sổ Edge dạng app tại `http://127.0.0.1:8620`.
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
