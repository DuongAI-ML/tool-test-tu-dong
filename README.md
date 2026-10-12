# Tool Test Tự Động

Nhập địa chỉ web → tool tự dò trang, sinh **Test Scenario** / **Test Case**, tự chạy test trên Edge / Chrome thật của máy và xuất Excel theo mẫu cột của công ty.

> Hướng dẫn sử dụng chi tiết: [HUONG-DAN.md](HUONG-DAN.md) · Dành cho người phát triển: [PHAT-TRIEN.md](PHAT-TRIEN.md)

## Yêu cầu máy

- Windows 10 hoặc 11 (64-bit).
- **Microsoft Edge** (Windows nào cũng có sẵn) hoặc Google Chrome.
- Không cần tài khoản, không gửi dữ liệu ra ngoài: tool chạy hoàn toàn trên máy.

## Cài đặt và chạy

Chọn **1 trong 2 cách**.

### Cách 1 — Bản EXE, không cần cài gì (khuyên dùng)

1. Vào **[Releases](https://github.com/DuongAI-ML/tool-test-tu-dong/releases/latest)**, tải file `ToolTestTuDong_v…_Windows.zip` (khoảng 57 MB).
2. Chuột phải file zip → **Extract All…** (Giải nén tất cả). Đừng chạy thẳng trong file zip.
3. Mở thư mục `ToolTestTuDong` → bấm đúp **`ToolTestTuDong.exe`**.
4. Muốn có biểu tượng ngoài Desktop: chuột phải `ToolTestTuDong.exe` → **Send to → Desktop (create shortcut)**.

Nếu Windows hiện **"Windows protected your PC"**: bấm **More info** → **Run anyway**. Tool chưa có chữ ký số nên Windows cảnh báo, không phải virus.

### Cách 2 — Chạy từ mã nguồn (tải nhẹ, dễ cập nhật)

1. Tải mã nguồn, chọn một trong hai:
   - Bấm nút xanh **Code → Download ZIP** ở đầu trang này, rồi giải nén.
   - Hoặc, nếu máy có git:
     ```bat
     git clone https://github.com/DuongAI-ML/tool-test-tu-dong.git
     ```
2. Mở thư mục vừa tải → bấm đúp **`ChayTool.bat`**.
   - **Lần đầu** (cần mạng, 1–3 phút): tool tự tìm Python. Nếu máy chưa có, tool tự cài Python 3.12 qua winget. Sau đó tool cài thư viện (khoảng 45 MB) vào `%LOCALAPPDATA%\ToolTestTuDong\venv`.
   - **Các lần sau**: bấm đúp `ChayTool.bat` là tool mở ngay.
3. Cập nhật bản mới: chạy `git pull` (hoặc tải ZIP mới), rồi bấm lại `ChayTool.bat`.

Không cần chạy `playwright install`: tool điều khiển Edge / Chrome có sẵn trên máy.

## Dùng thử lần đầu (2 phút)

1. Tool mở ra trong một cửa sổ Edge, ở tab **🌐 Test web**.
2. Bước 3 → bấm **"Dùng thử web demo"**, nhập mật khẩu demo `123456` → bấm **🔍 Dò trang**.
3. Bước 1 chọn **Test Scenario** hoặc **Test Case** → xem bảng ở bước 5.
4. Bước 6 → bấm **📄 Sinh file Excel mới** (hoặc **▶ Chạy test** nếu chọn Test Case).

Web demo có sẵn vài lỗi cố ý để bạn xem tool báo lỗi thế nào.

## Gặp sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| "Windows protected your PC" | **More info → Run anyway** |
| Bấm exe không thấy gì | Đợi 5–10 giây (lần đầu mở chậm hơn). Nếu vẫn không có gì: mở `http://127.0.0.1:8620` bằng Edge |
| `ChayTool.bat` báo không cài được thư viện | Kiểm tra mạng / proxy công ty, rồi chạy lại. Hoặc dùng **Cách 1** (bản EXE không cần mạng) |
| `ChayTool.bat` báo không có Python | Cài Python 3.12 tại https://www.python.org/downloads/ (tick **Add python.exe to PATH**), rồi chạy lại |
| "Không mở được trình duyệt" | Cài / cập nhật Microsoft Edge hoặc Google Chrome |
| Dò trang ra 0 nút / 0 form | Đọc khung ⚠ ngay dưới ô địa chỉ: tool ghi rõ lý do (cần đăng nhập, captcha, trang tải chậm…) và có nút xử lý |

## Dữ liệu và gỡ cài đặt

- Cài đặt, kết quả các lần chạy và phiên đăng nhập nằm ở `%LOCALAPPDATA%\ToolTestTuDong`. Tool **không lưu mật khẩu**.
- Gỡ cài đặt:
  1. Tắt tool (nút ⏻ góc trên bên phải).
  2. Xóa thư mục tool.
  3. Xóa thư mục `%LOCALAPPDATA%\ToolTestTuDong`.
